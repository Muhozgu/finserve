"""
generators/credit_history.py
-----------------------------
Generates CREDIT_HISTORY: periodic bureau-style snapshots per customer
(1 customer -> N records over time), as required by the spec.

Financial realism:
- credit_score is built from the same underlying customer "credit quality"
  latent factor used elsewhere (income, employment, tenure) plus a random
  walk over time (bureau scores drift slowly, they don't teleport).
- total_outstanding_debt and credit_utilization move together and are
  influenced by income (higher earners tend to carry higher absolute debt
  but not necessarily higher utilization).
- previous_defaults / previous_late_payments are monotonically
  non-decreasing counters per customer across snapshots (history doesn't
  erase itself), with occasional new defaults added at later snapshot dates.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from utils.helpers import make_ids


def _base_credit_quality(annual_income: np.ndarray, employment_status: np.ndarray,
                          employment_length: np.ndarray, rng: np.random.Generator) -> np.ndarray:
    """Latent 300-850 style baseline score before time-series drift is applied."""
    income_z = (np.log(np.maximum(annual_income, 1)) - np.log(30_000))
    status_adj = np.select(
        [employment_status == "Employed", employment_status == "Self-employed",
         employment_status == "Retired", employment_status == "Student",
         employment_status == "Unemployed"],
        [15, -5, 10, -10, -40],
        default=0,
    )
    tenure_adj = np.minimum(employment_length, 15) * 2.0
    noise = rng.normal(0, 35, size=len(annual_income))
    score = 650 + income_z * 40 + status_adj + tenure_adj + noise
    return np.clip(score, 300, 850)


def generate_credit_history(
    rng: np.random.Generator,
    n_target: int,
    customers: pd.DataFrame,
    history_start_year: int,
    as_of_date: str,
) -> pd.DataFrame:
    as_of = pd.Timestamp(as_of_date)
    n_customers = len(customers)

    base_score = _base_credit_quality(
        customers["annual_income"].values,
        customers["employment_status"].values,
        customers["employment_length_years"].values,
        rng,
    )
    base_score_by_cust = dict(zip(customers["customer_id"], base_score))
    income_by_cust = dict(zip(customers["customer_id"], customers["annual_income"]))

    # Decide how many snapshot records each customer gets, then trim/pad to
    # hit n_target overall (semi-annual-ish cadence since account creation).
    avg_records_per_customer = max(1, round(n_target / n_customers))
    records_per_customer = np.clip(
        rng.poisson(avg_records_per_customer, size=n_customers), 1, avg_records_per_customer * 3
    )

    # ---- "Wave" vectorization ----
    # A per-customer random walk (score drifts over time, defaults/late
    # payments accumulate) is inherently sequential *in time*, but customers
    # are independent of each other. So instead of looping over n_customers
    # (which can be 100k+) we loop over max_k (typically a handful of
    # snapshots per customer) and vectorize each "wave" across every
    # customer at once. This is what keeps this generator usable at the
    # 100k-customer / 500k+ row scale described in the spec.
    created = pd.to_datetime(customers["created_at"].values).values  # numpy datetime64[ns] array
    window_start = np.maximum(created, np.datetime64(f"{history_start_year}-01-01"))
    as_of_np = as_of.to_datetime64()
    total_days = np.asarray(np.maximum((as_of_np - window_start) / np.timedelta64(1, "D"), 0))

    max_k = int(records_per_customer.max())
    # Ragged per-customer offsets: draw max_k uniform offsets per customer,
    # scaled to that customer's own available window, then sort ascending
    # so snapshot dates move forward in time.
    unit_draws = rng.random(size=(n_customers, max_k))
    offsets = np.floor(unit_draws * total_days[:, None]).astype(int)
    offsets.sort(axis=1)
    snapshot_dates_matrix = window_start[:, None] + offsets.astype("timedelta64[D]")

    valid_mask = np.arange(max_k)[None, :] < records_per_customer[:, None]

    score_walk = base_score.copy()
    cum_defaults = np.zeros(n_customers, dtype=int)
    cum_late = np.zeros(n_customers, dtype=int)
    customer_id_arr = customers["customer_id"].values
    income_arr = customers["annual_income"].values

    wave_frames = []
    for c in range(max_k):
        score_walk = np.clip(score_walk + rng.normal(0, 12, size=n_customers), 300, 850)
        credit_score = np.round(score_walk).astype(int)

        score_z = (credit_score - 650) / 100
        utilization = np.clip(rng.beta(2, 3, size=n_customers) - score_z * 0.08, 0.0, 1.0)
        debt_mult = rng.uniform(0.2, 0.6, size=n_customers)
        total_outstanding_debt = np.round(np.maximum(0.0, income_arr * utilization * debt_mult), 2)

        open_accounts = np.maximum(0, rng.poisson(np.maximum(4 - score_z, 0.1)))
        closed_accounts = np.maximum(0, rng.poisson(2, size=n_customers))

        new_default = rng.random(n_customers) < np.clip(0.02 - score_z * 0.015, 0.001, 0.15)
        new_late = rng.random(n_customers) < np.clip(0.08 - score_z * 0.04, 0.005, 0.35)
        cum_defaults = cum_defaults + new_default.astype(int)
        cum_late = cum_late + new_late.astype(int) * rng.integers(1, 3, size=n_customers)

        col_mask = valid_mask[:, c]
        wave_frames.append(pd.DataFrame({
            "customer_id": customer_id_arr[col_mask],
            "record_date": pd.to_datetime(snapshot_dates_matrix[col_mask, c]).strftime("%Y-%m-%d"),
            "credit_score": credit_score[col_mask],
            "total_outstanding_debt": total_outstanding_debt[col_mask],
            "credit_utilization": np.round(utilization[col_mask], 4),
            "number_of_open_accounts": open_accounts[col_mask],
            "number_of_closed_accounts": closed_accounts[col_mask],
            "previous_defaults": cum_defaults[col_mask],
            "previous_late_payments": cum_late[col_mask],
        }))

    df = pd.concat(wave_frames, ignore_index=True)
    df = df.sort_values(["customer_id", "record_date"]).reset_index(drop=True)
    df.insert(0, "credit_history_id", [f"CRHIST{i:07d}" for i in range(1, len(df) + 1)])

    # Trim to n_target if the per-customer expansion overshot meaningfully
    if len(df) > n_target * 1.5:
        df = df.sample(n=n_target, random_state=rng.integers(0, 2**31 - 1)).reset_index(drop=True)

    return df

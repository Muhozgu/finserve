"""
generators/applications.py
---------------------------
Generates APPLICATIONS, a child of CUSTOMERS (1 customer -> N applications).

Financial realism:
- requested_amount scales with annual_income (people generally don't request
  loans wildly disproportionate to what they earn, though some do -- that
  variance is preserved).
- application_status probabilities are influenced lightly by income/DTI-like
  signal so that "Rejected" isn't purely random noise.
- decision_date is always >= application_date (this entity is generated
  *clean*; error_injection.py is responsible for later corrupting a small
  fraction of decision_date/application_date pairs on purpose).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from utils.helpers import make_ids, sample_indices


LOAN_PURPOSES = ["Debt Consolidation", "Home Improvement", "Auto", "Medical",
                  "Education", "Business", "Personal", "Wedding"]
CHANNELS = ["Online", "Branch", "Mobile App", "Call Center", "Partner Referral"]
CHANNEL_WEIGHTS = [0.42, 0.18, 0.25, 0.08, 0.07]

APPLICATION_STATUSES = ["Approved", "Rejected", "Pending", "Withdrawn"]


def generate_applications(
    rng: np.random.Generator,
    n: int,
    customers: pd.DataFrame,
    as_of_date: str,
) -> pd.DataFrame:
    n_customers = len(customers)
    # Sample which customer each application belongs to (with replacement --
    # customers can have multiple applications, matching the 1:N spec).
    cust_positions = sample_indices(rng, n_customers, n, replace=True)
    customer_id = customers["customer_id"].values[cust_positions]
    annual_income = customers["annual_income"].values[cust_positions]

    application_id = make_ids("APP", n)

    # Requested amount: roughly 0.1x - 1.2x annual income, lognormal-ish spread
    income_multiplier = np.clip(rng.gamma(shape=2.0, scale=0.25, size=n), 0.05, 2.0)
    requested_amount = np.round(np.maximum(annual_income * income_multiplier, 500), 2)

    loan_purpose = rng.choice(LOAN_PURPOSES, size=n)
    application_channel = rng.choice(CHANNELS, size=n, p=CHANNEL_WEIGHTS)

    # application_date spread over the last ~5 years up to as_of_date
    as_of = pd.Timestamp(as_of_date)
    start = as_of - pd.Timedelta(days=5 * 365)
    offsets = rng.integers(0, (as_of - start).days + 1, size=n)
    application_date = pd.to_datetime(start) + pd.to_timedelta(offsets, unit="D")

    # Status probability lightly informed by requested_amount vs income (higher
    # ratio -> somewhat higher rejection chance) -- keeps things non-trivial
    # for anyone building a rejection-rate dashboard later.
    ratio = requested_amount / np.maximum(annual_income, 1)
    reject_bias = np.clip(ratio / 2, 0, 0.35)
    status = np.empty(n, dtype=object)
    for i in range(n):
        p_reject = 0.20 + reject_bias[i]
        p_reject = min(p_reject, 0.6)
        remaining = 1 - p_reject
        probs = [remaining * 0.68, p_reject, remaining * 0.20, remaining * 0.12]
        probs = np.array(probs)
        probs = probs / probs.sum()
        status[i] = rng.choice(APPLICATION_STATUSES, p=probs)

    # decision_date: a handful of days after application_date, only for
    # non-pending applications (pending realistically has no decision yet)
    decision_lag_days = rng.integers(1, 21, size=n)
    decision_date = application_date + pd.to_timedelta(decision_lag_days, unit="D")
    decision_date = decision_date.where(pd.Series(status) != "Pending", pd.NaT)
    # Clip decision dates that would land in the future relative to as_of
    decision_date = decision_date.where(decision_date <= as_of, as_of)

    df = pd.DataFrame({
        "application_id": application_id,
        "customer_id": customer_id,
        "application_date": application_date.strftime("%Y-%m-%d"),
        "requested_amount": requested_amount,
        "loan_purpose": loan_purpose,
        "application_status": status,
        "application_channel": application_channel,
        "decision_date": pd.Series(decision_date).dt.strftime("%Y-%m-%d"),
    })
    return df

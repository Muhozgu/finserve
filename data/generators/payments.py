"""
generators/payments.py
-----------------------
Generates PAYMENTS, a child of LOANS (1 loan -> N payments).

Financial realism:
- due_date schedule is built off the loan's start_date and term_months at a
  monthly cadence, up to as_of_date (a loan doesn't have payments due in the
  future yet from the simulated "today").
- payment behavior (on-time / late / missed / partial) is driven by a
  per-loan "payment reliability" draw so that a given loan's payments are
  correlated with each other (a customer who misses one payment is more
  likely to miss others) rather than each payment being independently random.
- days_late and payment_status are derived consistently from payment_date vs
  due_date.
- Defaulted/Delinquent loans get a worse reliability draw than Active/Paid Off
  loans, so loan_status and payment behavior agree with each other in the
  clean data (error_injection.py later breaks this consistency for a small
  subset, per your "defaulted loan has suspicious payment status" example).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from data.utils.helpers import make_ids


def generate_payments(
    rng: np.random.Generator,
    n_target: int,
    loans: pd.DataFrame,
    as_of_date: str,
) -> pd.DataFrame:
    as_of = pd.Timestamp(as_of_date)

    # Build the full "schedule" of due payments per loan, fully vectorized
    # (no per-loan Python loop -- this matters a lot at 500k+ loans / 1M+
    # payments; a naive row-by-row loop here was the single biggest
    # performance bottleneck measured during development).
    starts = pd.to_datetime(loans["start_date"]).values
    terms = loans["term_months"].values.astype(int)
    statuses = loans["loan_status"].values
    loan_ids = loans["loan_id"].values
    monthly_payments = loans["monthly_payment"].values

    # Per-loan reliability: 0 = always pays on time, 1 = very unreliable.
    # Defaulted/Delinquent loans skew unreliable; Paid Off/Active skew reliable.
    base_reliability = rng.beta(1.5, 6.0, size=len(loans))  # skewed low (mostly reliable)
    status_bump = np.select(
        [statuses == "Defaulted", statuses == "Delinquent", statuses == "Active", statuses == "Paid Off"],
        [0.55, 0.30, 0.0, -0.05],
        default=0.0,
    )
    reliability = np.clip(base_reliability + status_bump, 0.01, 0.95)

    # How many monthly payments have plausibly come due by as_of_date for
    # each loan (approximating a month as 30.44 days, consistent with how
    # maturity_date is derived in loans.py), capped at the loan's term.
    days_elapsed = (as_of.to_numpy() - starts) / np.timedelta64(1, "D")
    months_elapsed = np.floor(days_elapsed / 30.44).astype(int)
    n_due = np.clip(np.minimum(months_elapsed, terms), 0, None)

    total_rows = int(n_due.sum())
    if total_rows == 0:
        return pd.DataFrame(columns=[
            "payment_id", "loan_id", "customer_id", "due_date", "payment_date",
            "amount_due", "amount_paid", "days_late", "payment_status",
        ])

    # Vectorized "ragged range": for each loan i with n_due[i] payments,
    # produce payment_number = 1..n_due[i]. Standard NumPy trick using
    # repeat + cumulative offsets instead of a Python loop per loan.
    row_loan_idx = np.repeat(np.arange(len(loans)), n_due)
    cum_n_due = np.cumsum(n_due)
    start_of_group = np.repeat(cum_n_due - n_due, n_due)
    payment_number = (np.arange(total_rows) - start_of_group) + 1

    due_date = (
        pd.to_datetime(starts[row_loan_idx])
        + pd.to_timedelta(np.round(payment_number * 30.44).astype(int), unit="D")
    )

    sched = pd.DataFrame({
        "loan_id": loan_ids[row_loan_idx],
        "due_date": due_date,
        "amount_due": monthly_payments[row_loan_idx],
        "reliability": reliability[row_loan_idx],
    })

    # Subsample to n_target if we generated more than requested; if fewer,
    # keep all (a full realistic schedule may simply be smaller than the ask
    # for small configs -- documented in README performance/config notes).
    if len(sched) > n_target:
        sched = sched.sample(n=n_target, random_state=rng.integers(0, 2**31 - 1)).reset_index(drop=True)
    else:
        sched = sched.reset_index(drop=True)

    n = len(sched)
    roll = rng.random(n)
    rel = sched["reliability"].values

    # Outcome buckets, probability driven by that loan's reliability score
    p_missed = 0.02 + rel * 0.35
    p_late = 0.05 + rel * 0.30
    p_partial = 0.02 + rel * 0.10
    outcome = np.where(
        roll < p_missed, "Missed",
        np.where(roll < p_missed + p_late, "Late",
                 np.where(roll < p_missed + p_late + p_partial, "Partial", "On-time"))
    )

    days_late = np.zeros(n, dtype=int)
    late_mask = outcome == "Late"
    days_late[late_mask] = rng.integers(1, 45, size=late_mask.sum())
    missed_mask = outcome == "Missed"
    days_late[missed_mask] = rng.integers(45, 120, size=missed_mask.sum())

    payment_date = pd.to_datetime(sched["due_date"]) + pd.to_timedelta(days_late, unit="D")
    # On-time / Partial payments happen on or slightly before/at the due date
    ontime_or_partial = ~(late_mask | missed_mask)
    jitter = rng.integers(-2, 1, size=ontime_or_partial.sum())
    payment_date_vals = payment_date.values.copy()
    payment_date_vals[ontime_or_partial] = (
        pd.to_datetime(sched["due_date"])[ontime_or_partial].values + pd.to_timedelta(jitter, unit="D")
    )
    payment_date = pd.Series(payment_date_vals)
    # Missed payments realistically have no payment_date at all (never paid)
    payment_date = payment_date.where(~missed_mask, pd.NaT)

    amount_due = sched["amount_due"].values
    amount_paid = np.where(
        outcome == "Partial", np.round(amount_due * rng.uniform(0.3, 0.85, size=n), 2),
        np.where(outcome == "Missed", 0.0, amount_due),
    )

    payment_id = make_ids("PAY", n)

    df = pd.DataFrame({
        "payment_id": payment_id,
        "loan_id": sched["loan_id"].values,
        "due_date": pd.to_datetime(sched["due_date"]).dt.strftime("%Y-%m-%d"),
        "payment_date": payment_date.dt.strftime("%Y-%m-%d"),
        "amount_due": amount_due,
        "amount_paid": amount_paid,
        "days_late": np.where(missed_mask, np.nan, days_late),
        "payment_status": outcome,
    })

    # attach customer_id via loans lookup (denormalized FK for convenience,
    # same pattern the spec uses on payments)
    loan_to_cust = dict(zip(loans["loan_id"], loans["customer_id"]))
    df["customer_id"] = df["loan_id"].map(loan_to_cust)
    df = df[["payment_id", "loan_id", "customer_id", "due_date", "payment_date",
             "amount_due", "amount_paid", "days_late", "payment_status"]]
    return df

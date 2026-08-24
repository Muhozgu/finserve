"""
generators/loans.py
--------------------
Generates LOANS. Per the spec, an application has a 0/1 relationship to a
loan -- only APPROVED applications ever originate a loan.

Financial realism:
- loan_amount is derived from (and <=) the application's requested_amount,
  with a small negotiation haircut applied to some loans.
- interest_rate is priced off a lightweight internal "underwriting risk
  proxy" built from the customer's income and employment status (NOT the
  same as the later, richer risk_assessments.probability_of_default, which
  intentionally uses more signals including credit_history -- see README
  section "How realistic financial relationships are generated" for why
  these are deliberately not the same formula).
- monthly_payment is computed with the standard amortization formula from
  loan_amount, interest_rate, and term_months -- these three are always
  mathematically consistent for cleanly-generated rows (error_injection.py
  may later break this on purpose for a small subset).
- loan_status is seeded here as a plausible starting point and is later
  reconciled against payment behavior in payments.py logic notes (a full
  payment-driven status recompute is intentionally left as a dbt exercise --
  see README).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from utils.helpers import make_ids


TERM_OPTIONS = [12, 24, 36, 48, 60, 72, 84]
TERM_WEIGHTS = [0.10, 0.18, 0.22, 0.18, 0.16, 0.10, 0.06]

LOAN_STATUSES = ["Active", "Paid Off", "Defaulted", "Delinquent"]


def _underwriting_risk_proxy(annual_income: np.ndarray, employment_status: np.ndarray,
                              employment_length: np.ndarray, rng: np.random.Generator) -> np.ndarray:
    """
    A simple 0-1 risk proxy used ONLY for pricing (interest rate) at
    origination time, since real credit_history records haven't been
    "pulled" yet in this simplified simulation. Higher = riskier.
    """
    income_z = (np.log(np.maximum(annual_income, 1)) - np.log(30_000)) / 1.0
    income_component = -income_z  # higher income -> lower risk

    status_risk = np.select(
        [employment_status == "Employed", employment_status == "Self-employed",
         employment_status == "Retired", employment_status == "Student",
         employment_status == "Unemployed"],
        [-0.3, 0.1, -0.1, 0.2, 0.9],
        default=0.0,
    )

    tenure_component = -np.minimum(employment_length, 15) / 15.0

    noise = rng.normal(0, 0.5, size=len(annual_income))
    raw = income_component * 0.5 + status_risk + tenure_component * 0.4 + noise
    # squash to 0-1
    return 1 / (1 + np.exp(-raw))


def generate_loans(
    rng: np.random.Generator,
    n_target: int,
    applications: pd.DataFrame,
    customers: pd.DataFrame,
    as_of_date: str,
) -> pd.DataFrame:
    approved = applications[applications["application_status"] == "Approved"].copy()
    n = min(n_target, len(approved))
    if n < n_target:
        # Not a hard failure -- just means the requested n_loans exceeded the
        # number of approved applications produced. We cap gracefully and the
        # orchestrator will log this to stdout.
        pass
    approved = approved.sample(n=n, random_state=rng.integers(0, 2**31 - 1))

    cust_lookup = customers.set_index("customer_id")
    cust_income = cust_lookup.loc[approved["customer_id"], "annual_income"].values
    cust_status = cust_lookup.loc[approved["customer_id"], "employment_status"].values
    cust_tenure = cust_lookup.loc[approved["customer_id"], "employment_length_years"].values

    loan_id = make_ids("LOAN", n)

    # Negotiation: most loans equal requested_amount; some get a haircut
    haircut = rng.uniform(0.7, 1.0, size=n)
    apply_haircut = rng.random(n) < 0.35
    loan_amount = np.where(apply_haircut, approved["requested_amount"].values * haircut,
                            approved["requested_amount"].values)
    loan_amount = np.round(np.maximum(loan_amount, 500), 2)

    risk_proxy = _underwriting_risk_proxy(cust_income, cust_status, cust_tenure, rng)
    # Base rate + risk premium + small noise, floored/capped to realistic consumer-loan range
    interest_rate = 3.5 + risk_proxy * 14.0 + rng.normal(0, 0.6, size=n)
    interest_rate = np.round(np.clip(interest_rate, 2.5, 24.9), 2)

    term_months = rng.choice(TERM_OPTIONS, size=n, p=TERM_WEIGHTS)

    # start_date: shortly after the application's decision_date
    decision_date = pd.to_datetime(approved["decision_date"].values)
    lag_days = rng.integers(1, 15, size=n)
    start_date = decision_date + pd.to_timedelta(lag_days, unit="D")
    maturity_date = start_date + pd.to_timedelta((term_months * 30.44).astype(int), unit="D")

    # Standard amortization formula: M = P * r(1+r)^n / ((1+r)^n - 1), r = monthly rate
    monthly_rate = interest_rate / 100 / 12
    n_payments = term_months
    with np.errstate(divide="ignore", invalid="ignore"):
        factor = (1 + monthly_rate) ** n_payments
        monthly_payment = loan_amount * monthly_rate * factor / (factor - 1)
    monthly_payment = np.round(monthly_payment, 2)

    # loan_status: seeded plausibly based on how far along the term we are
    # relative to as_of_date, plus a risk-weighted default/delinquency chance.
    as_of = pd.Timestamp(as_of_date)
    matured = (maturity_date <= as_of)
    default_prob = 0.03 + risk_proxy * 0.15
    delinquent_prob = 0.04 + risk_proxy * 0.10
    roll = rng.random(n)
    status = np.where(
        matured, "Paid Off",
        np.where(roll < default_prob, "Defaulted",
                 np.where(roll < default_prob + delinquent_prob, "Delinquent", "Active"))
    )

    df = pd.DataFrame({
        "loan_id": loan_id,
        "application_id": approved["application_id"].values,
        "customer_id": approved["customer_id"].values,
        "loan_amount": loan_amount,
        "interest_rate": interest_rate,
        "term_months": term_months,
        "start_date": start_date.strftime("%Y-%m-%d"),
        "maturity_date": maturity_date.strftime("%Y-%m-%d"),
        "loan_status": status,
        "monthly_payment": monthly_payment,
    })
    return df

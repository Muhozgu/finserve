"""
generators/risk_assessments.py
--------------------------------
Generates RISK_ASSESSMENTS: one row per (customer, loan) for loans, plus
one row per (customer, NULL loan_id) for rejected applications -- because in
a real credit-risk engine you score BEFORE you decide whether a loan exists.

This is the core "credit risk logic" module the spec asks for. See the
formula documented inline and in the README's "Credit Risk Logic" section.

IMPORTANT: This is a synthetic, illustrative scoring formula built for a
portfolio project. It is NOT a real bank's proprietary risk model and should
never be represented as one.

probability_of_default (PD) formula
------------------------------------
We standardize (z-score) each risk driver, apply a hand-picked weight whose
SIGN matches real-world credit risk intuition (see table below), sum them
into a single "risk index", add Gaussian noise for realistic variation, and
squash the result through a logistic function to land in (0, 1).

    risk_index = 0.9*z(-credit_score)
               + 0.8*z(debt_to_income_ratio)
               + 0.6*z(credit_utilization)
               + 0.7*z(previous_defaults)
               + 0.4*z(previous_late_payments)
               + 0.3*z(-employment_length_years)
               + 0.5*z(loan_to_income_ratio)
               + 0.2*z(loan_term_months)
               + 0.3*z(interest_rate)
               + noise ~ Normal(0, 0.5)

    PD = sigmoid(risk_index * 0.6 - 2.4)   # scale/shift tuned so average PD sits ~10-13%, median ~8%

risk_score: rescaled inverse of PD onto a familiar 300-850 band, so it reads
like a bureau-style score (higher = safer), purely for dashboard usability.

risk_category: LOW / MEDIUM / HIGH, cut at PD < 0.10 and PD < 0.30, with a
little category jitter (a few points near each boundary can land either
side) so the boundary isn't perfectly clean -- deliberately gives you
something to sanity-check with a dbt test later.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from data_source.utils.helpers import make_ids, zscore, logistic


def _compute_pd(
    credit_score, dti, utilization, prev_defaults, prev_late,
    employment_length, loan_to_income, loan_term, interest_rate, rng,
):
    n = len(credit_score)
    z = lambda arr: zscore(np.asarray(arr, dtype=float))

    risk_index = (
        0.9 * z(-np.asarray(credit_score, dtype=float))
        + 0.8 * z(dti)
        + 0.6 * z(utilization)
        + 0.7 * z(prev_defaults)
        + 0.4 * z(prev_late)
        + 0.3 * z(-np.asarray(employment_length, dtype=float))
        + 0.5 * z(loan_to_income)
        + 0.2 * z(loan_term)
        + 0.3 * z(interest_rate)
        + rng.normal(0, 0.5, size=n)
    )
    pd_values = logistic(risk_index * 0.6 - 2.4)
    return np.clip(pd_values, 0.001, 0.999)


def _pd_to_risk_score(pd_values: np.ndarray) -> np.ndarray:
    # Map PD (0-1) inversely onto a 300-850 band via logit transform, then rescale.
    pd_clipped = np.clip(pd_values, 0.001, 0.999)
    logit = np.log(pd_clipped / (1 - pd_clipped))
    # logit roughly ranges -6..6 in practice; map -6->850, 6->300 linearly
    score = 850 - (logit + 6) / 12 * (850 - 300)
    return np.clip(score, 300, 850).round().astype(int)


def _risk_category(pd_values: np.ndarray, rng: np.random.Generator) -> np.ndarray:
    jitter = rng.normal(0, 0.02, size=len(pd_values))
    jittered = pd_values + jitter
    return np.select(
        [jittered < 0.10, jittered < 0.30],
        ["LOW", "MEDIUM"],
        default="HIGH",
    )


def generate_risk_assessments(
    rng: np.random.Generator,
    customers: pd.DataFrame,
    loans: pd.DataFrame,
    applications: pd.DataFrame,
    credit_history: pd.DataFrame,
    as_of_date: str,
    model_version: str,
) -> pd.DataFrame:
    as_of = pd.Timestamp(as_of_date)

    # Latest credit_history snapshot per customer (most recent record_date)
    if len(credit_history) > 0:
        ch_sorted = credit_history.sort_values("record_date")
        latest_ch = ch_sorted.groupby("customer_id").tail(1).set_index("customer_id")
    else:
        latest_ch = pd.DataFrame(columns=[
            "credit_score", "total_outstanding_debt", "credit_utilization",
            "previous_defaults", "previous_late_payments",
        ]).set_index(pd.Index([], name="customer_id"))

    cust_lookup = customers.set_index("customer_id")

    def _lookup_credit_fields(customer_ids: pd.Series):
        """Pull latest credit_history fields per customer_id, falling back to
        population-median defaults for customers with no history yet."""
        med_score = int(latest_ch["credit_score"].median()) if len(latest_ch) else 650
        med_debt_ratio = 0.25

        idx = customer_ids.values
        has_ch = pd.Series(idx).isin(latest_ch.index).values

        credit_score = np.full(len(idx), med_score, dtype=float)
        utilization = np.full(len(idx), med_debt_ratio, dtype=float)
        outstanding_debt = np.zeros(len(idx), dtype=float)
        prev_defaults = np.zeros(len(idx), dtype=float)
        prev_late = np.zeros(len(idx), dtype=float)

        if has_ch.any():
            matched = latest_ch.loc[idx[has_ch]]
            credit_score[has_ch] = matched["credit_score"].values
            utilization[has_ch] = matched["credit_utilization"].values
            outstanding_debt[has_ch] = matched["total_outstanding_debt"].values
            prev_defaults[has_ch] = matched["previous_defaults"].values
            prev_late[has_ch] = matched["previous_late_payments"].values

        return credit_score, utilization, outstanding_debt, prev_defaults, prev_late

    # ---- 1. Loan-linked assessments ----
    loan_customer_ids = loans["customer_id"]
    credit_score, utilization, outstanding_debt, prev_defaults, prev_late = _lookup_credit_fields(loan_customer_ids)
    income = cust_lookup.loc[loan_customer_ids, "annual_income"].values
    employment_length = cust_lookup.loc[loan_customer_ids, "employment_length_years"].values
    dti = outstanding_debt / np.maximum(income, 1)
    loan_to_income = loans["loan_amount"].values / np.maximum(income, 1)

    pd_loans = _compute_pd(
        credit_score, dti, utilization, prev_defaults, prev_late,
        employment_length, loan_to_income, loans["term_months"].values,
        loans["interest_rate"].values, rng,
    )
    risk_score_loans = _pd_to_risk_score(pd_loans)
    risk_category_loans = _risk_category(pd_loans, rng)

    # assessment_date: shortly after loan start (underwriting decision point)
    assessment_date_loans = pd.to_datetime(loans["start_date"]) - pd.to_timedelta(rng.integers(0, 4, size=len(loans)), unit="D")

    loans_assess = pd.DataFrame({
        "assessment_id": make_ids("RISK", len(loans)),
        "customer_id": loans["customer_id"].values,
        "loan_id": loans["loan_id"].values,
        "assessment_date": assessment_date_loans.dt.strftime("%Y-%m-%d"),
        "risk_score": risk_score_loans,
        "probability_of_default": np.round(pd_loans, 4),
        "risk_category": risk_category_loans,
        "model_version": model_version,
    })

    # ---- 2. Application-only assessments for rejected applications (no loan exists) ----
    rejected = applications[applications["application_status"] == "Rejected"].copy()
    if len(rejected) > 0:
        credit_score_r, utilization_r, outstanding_debt_r, prev_defaults_r, prev_late_r = _lookup_credit_fields(rejected["customer_id"])
        income_r = cust_lookup.loc[rejected["customer_id"], "annual_income"].values
        employment_length_r = cust_lookup.loc[rejected["customer_id"], "employment_length_years"].values
        dti_r = outstanding_debt_r / np.maximum(income_r, 1)
        loan_to_income_r = rejected["requested_amount"].values / np.maximum(income_r, 1)
        # No loan term/rate exists yet -- use population medians as neutral placeholders
        median_term = loans["term_months"].median() if len(loans) else 36
        median_rate = loans["interest_rate"].median() if len(loans) else 12.0

        pd_rejected = _compute_pd(
            credit_score_r, dti_r, utilization_r, prev_defaults_r, prev_late_r,
            employment_length_r, loan_to_income_r,
            np.full(len(rejected), median_term), np.full(len(rejected), median_rate), rng,
        )
        risk_score_rejected = _pd_to_risk_score(pd_rejected)
        risk_category_rejected = _risk_category(pd_rejected, rng)

        rejected_assess = pd.DataFrame({
            "assessment_id": make_ids("RISKA", len(rejected)),
            "customer_id": rejected["customer_id"].values,
            "loan_id": None,
            "assessment_date": rejected["decision_date"].values,
            "risk_score": risk_score_rejected,
            "probability_of_default": np.round(pd_rejected, 4),
            "risk_category": risk_category_rejected,
            "model_version": model_version,
        })
        result = pd.concat([loans_assess, rejected_assess], ignore_index=True)
    else:
        result = loans_assess

    return result

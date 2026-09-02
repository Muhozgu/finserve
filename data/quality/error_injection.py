"""
quality/error_injection.py
----------------------------
The error-injection engine. Everything in generators/ produces CLEAN,
relationally-consistent data. This module is the ONLY place that
deliberately corrupts it, and every single corruption it makes is logged
to `data_quality_issues.csv` with a stable issue_id, so you can measure
recall/precision of your dbt tests against a ground truth.

Design principles
------------------
1. Never touch more than `rate` fraction of ELIGIBLE rows for a given error
   type (eligible = rows where that error type is applicable/meaningful).
2. Corruption happens on a COPY of the row set chosen by `rate`; we never
   let one row get "double corrupted" into something contradictory within
   the same error type (though a row CAN receive multiple different error
   types -- that's realistic: a messy record is often messy in more than
   one way).
3. Every injected issue is logged via `IssueLog.add(...)` before/while the
   corruption is applied, so the log and the data can never drift apart.
4. Functions are pure-ish: they take a DataFrame and return a
   (possibly-longer, for duplicates) corrupted DataFrame. The caller
   (generate_data.py) decides the order in which error types are applied
   per entity.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Dict, Any

import numpy as np
import pandas as pd

from data_source.config import (
    EMPLOYMENT_STATUS_VARIANTS,
    APPLICATION_STATUS_VARIANTS,
    PAYMENT_STATUS_VARIANTS,
)


@dataclass
class IssueLog:
    """Accumulates data_quality_issues.csv rows across the whole run."""
    rows: List[Dict[str, Any]] = field(default_factory=list)
    _counter: int = 0
    injected_at: str = "2026-08-24"

    def add(self, entity: str, record_id, issue_type: str, description: str, severity: str):
        self._counter += 1
        self.rows.append({
            "issue_id": self._counter,
            "entity": entity,
            "record_id": record_id,
            "issue_type": issue_type,
            "issue_description": description,
            "severity": severity,
            "injected_at": self.injected_at,
        })

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame(self.rows, columns=[
            "issue_id", "entity", "record_id", "issue_type",
            "issue_description", "severity", "injected_at",
        ])


def _pick_rows(rng: np.random.Generator, n: int, rate: float) -> np.ndarray:
    """Return boolean mask selecting ~rate fraction of n rows."""
    if n == 0 or rate <= 0:
        return np.zeros(n, dtype=bool)
    k = max(0 if rate == 0 else 1, int(round(n * rate))) if rate > 0 else 0
    k = min(k, n)
    idx = rng.choice(n, size=k, replace=False)
    mask = np.zeros(n, dtype=bool)
    mask[idx] = True
    return mask


def _id_col(entity: str) -> str:
    return {
        "customers": "customer_id",
        "applications": "application_id",
        "loans": "loan_id",
        "payments": "payment_id",
        "credit_history": "credit_history_id",
        "risk_assessments": "assessment_id",
    }[entity]


# ---------------------------------------------------------------------------
# A. Missing values
# ---------------------------------------------------------------------------
def inject_missing_values(df: pd.DataFrame, entity: str, columns: List[str], rate: float,
                           rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = df.copy()
    id_col = _id_col(entity)
    for col in columns:
        if col not in df.columns:
            continue
        mask = _pick_rows(rng, len(df), rate)
        for i in np.where(mask)[0]:
            log.add(entity, df.iloc[i][id_col], f"missing_{col}",
                     f"{col} was set to NULL (simulated missing operational data)", "medium")
        df.loc[mask, col] = np.nan
    return df


# ---------------------------------------------------------------------------
# B. Duplicate records
# ---------------------------------------------------------------------------
def inject_duplicates(df: pd.DataFrame, entity: str, rate: float,
                       rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    if len(df) == 0 or rate <= 0:
        return df
    id_col = _id_col(entity)
    mask = _pick_rows(rng, len(df), rate)
    dup_rows = df[mask].copy()
    for rid in dup_rows[id_col]:
        log.add(entity, rid, "duplicate_record",
                 f"Row with {id_col}={rid} was duplicated verbatim (simulated double-submit / re-sync)", "low")
    return pd.concat([df, dup_rows], ignore_index=True)


# ---------------------------------------------------------------------------
# C. Invalid numerical values
# ---------------------------------------------------------------------------
def inject_invalid_numeric(df: pd.DataFrame, entity: str, column: str, rate: float,
                            rng: np.random.Generator, log: IssueLog, mode: str = "negative",
                            bound: float = None) -> pd.DataFrame:
    """
    mode="negative": flips sign to negative.
    mode="out_of_range": sets value beyond a hard bound (e.g. credit_score > 850 or < 300).
    """
    df = df.copy()
    if column not in df.columns or len(df) == 0:
        return df
    id_col = _id_col(entity)
    mask = _pick_rows(rng, len(df), rate)
    for i in np.where(mask)[0]:
        rid = df.iloc[i][id_col]
        old_val = df.iloc[i][column]
        if mode == "negative":
            new_val = -abs(old_val) if pd.notna(old_val) and old_val != 0 else -abs(rng.uniform(100, 1000))
            log.add(entity, rid, f"negative_{column}",
                     f"{column} set to negative value ({new_val}) -- physically invalid", "high")
        elif mode == "out_of_range":
            new_val = bound + rng.uniform(1, 50) if rng.random() < 0.5 else -(rng.uniform(1, 50))
            new_val = round(new_val) if column == "credit_score" else new_val
            log.add(entity, rid, f"out_of_range_{column}",
                     f"{column} set to {new_val}, outside the valid domain (bound={bound})", "high")
        else:
            continue
        df.at[df.index[i], column] = new_val
    return df


def inject_impossible_age(customers: pd.DataFrame, rate: float, rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = customers.copy()
    mask = _pick_rows(rng, len(df), rate)
    for i in np.where(mask)[0]:
        rid = df.iloc[i]["customer_id"]
        # Impossible DOB: either implies age > 120 or a future birth date
        if rng.random() < 0.5:
            bad_dob = pd.Timestamp("1890-01-01") + pd.to_timedelta(int(rng.integers(0, 3650)), unit="D")
            desc = "date_of_birth implies an age over 120 years"
        else:
            bad_dob = pd.Timestamp("2026-08-24") + pd.to_timedelta(int(rng.integers(30, 3650)), unit="D")
            desc = "date_of_birth is set in the future"
        df.at[df.index[i], "date_of_birth"] = bad_dob.strftime("%Y-%m-%d")
        log.add("customers", rid, "impossible_age", desc, "high")
    return df


# ---------------------------------------------------------------------------
# D. Inconsistent categorical values
# ---------------------------------------------------------------------------
def inject_inconsistent_categories(df: pd.DataFrame, entity: str, column: str,
                                    variants_map: dict, rate: float,
                                    rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = df.copy()
    if column not in df.columns or len(df) == 0:
        return df
    id_col = _id_col(entity)
    mask = _pick_rows(rng, len(df), rate)
    for i in np.where(mask)[0]:
        canonical = df.iloc[i][column]
        if canonical not in variants_map:
            continue
        variants = variants_map[canonical]
        new_val = rng.choice(variants)
        rid = df.iloc[i][id_col]
        df.at[df.index[i], column] = new_val
        log.add(entity, rid, f"inconsistent_{column}",
                 f"{column} written as '{new_val}' instead of canonical '{canonical}'", "low")
    return df


# ---------------------------------------------------------------------------
# E. Invalid dates
# ---------------------------------------------------------------------------
def inject_invalid_dates_applications(applications: pd.DataFrame, rate: float,
                                       rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = applications.copy()
    eligible = df["decision_date"].notna()
    eligible_idx = np.where(eligible.values)[0]
    if len(eligible_idx) == 0:
        return df
    n_pick = max(1, int(round(len(eligible_idx) * rate))) if rate > 0 else 0
    picked = rng.choice(eligible_idx, size=min(n_pick, len(eligible_idx)), replace=False)
    for i in picked:
        rid = df.iloc[i]["application_id"]
        app_date = pd.Timestamp(df.iloc[i]["application_date"])
        bad_decision = app_date - pd.Timedelta(days=int(rng.integers(1, 30)))
        df.at[df.index[i], "decision_date"] = bad_decision.strftime("%Y-%m-%d")
        log.add("applications", rid, "decision_before_application",
                "decision_date falls before application_date", "high")
    return df


def inject_invalid_dates_loans(loans: pd.DataFrame, rate: float,
                                rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = loans.copy()
    mask = _pick_rows(rng, len(df), rate)
    for i in np.where(mask)[0]:
        rid = df.iloc[i]["loan_id"]
        start = pd.Timestamp(df.iloc[i]["start_date"])
        bad_maturity = start - pd.Timedelta(days=int(rng.integers(1, 60)))
        df.at[df.index[i], "maturity_date"] = bad_maturity.strftime("%Y-%m-%d")
        log.add("loans", rid, "maturity_before_start",
                "maturity_date falls before start_date", "high")
    return df


def inject_invalid_dates_payments(payments: pd.DataFrame, loans: pd.DataFrame, rate: float,
                                   rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = payments.copy()
    mask = _pick_rows(rng, len(df), rate)
    loan_start = dict(zip(loans["loan_id"], loans["start_date"]))
    for i in np.where(mask)[0]:
        rid = df.iloc[i]["payment_id"]
        loan_id = df.iloc[i]["loan_id"]
        start = loan_start.get(loan_id)
        if start is None or pd.isna(df.iloc[i]["payment_date"]):
            continue
        bad_payment_date = pd.Timestamp(start) - pd.Timedelta(days=int(rng.integers(1, 30)))
        df.at[df.index[i], "payment_date"] = bad_payment_date.strftime("%Y-%m-%d")
        log.add("payments", rid, "payment_before_loan_start",
                "payment_date falls before the loan's start_date", "high")
    return df


# ---------------------------------------------------------------------------
# F. Broken relationships (orphan foreign keys)
# ---------------------------------------------------------------------------
def inject_orphan_fk(df: pd.DataFrame, entity: str, fk_column: str, id_prefix: str,
                      rate: float, rng: np.random.Generator, log: IssueLog) -> pd.DataFrame:
    df = df.copy()
    id_col = _id_col(entity)
    mask = _pick_rows(rng, len(df), rate)
    for i in np.where(mask)[0]:
        rid = df.iloc[i][id_col]
        fake_fk = f"{id_prefix}{rng.integers(9_000_000, 9_999_999)}"
        df.at[df.index[i], fk_column] = fake_fk
        log.add(entity, rid, f"orphan_{fk_column}",
                 f"{fk_column} references '{fake_fk}', which does not exist in the parent table", "high")
    return df


# ---------------------------------------------------------------------------
# G. Business rule violations
# ---------------------------------------------------------------------------
def inject_business_rule_violations(loans: pd.DataFrame, applications: pd.DataFrame,
                                     payments: pd.DataFrame, rate: float,
                                     rng: np.random.Generator, log: IssueLog):
    loans = loans.copy()
    applications = applications.copy()
    payments = payments.copy()

    # G1: a rejected application ends up with an active loan attached
    #     (simulate by flipping a small number of Rejected -> having a loan
    #     already exists for that application_id via loans table -- since
    #     loans are only generated from Approved apps, we instead flip the
    #     application's own status label to Rejected AFTER the loan was made,
    #     which is exactly how this bug happens in real systems: someone
    #     corrects the decision after the fact without cascading it.)
    app_ids_with_loans = set(loans["application_id"])
    eligible = applications[
        (applications["application_status"] == "Approved") & (applications["application_id"].isin(app_ids_with_loans))
    ]
    if len(eligible) > 0:
        n_pick = max(1, int(round(len(eligible) * rate)))
        picked = eligible.sample(n=min(n_pick, len(eligible)), random_state=rng.integers(0, 2**31 - 1))
        applications.loc[picked.index, "application_status"] = "Rejected"
        for aid in picked["application_id"]:
            log.add("applications", aid, "rejected_with_active_loan",
                    "Application status is Rejected but a loan already exists for it", "high")

    # G2: payment amount negative (business-rule level, distinct from the
    #     generic negative-numeric injector -- flagged specifically here since
    #     a negative payment amount is a business rule break, not just a
    #     type/domain error)
    mask = _pick_rows(rng, len(payments), rate)
    for i in np.where(mask)[0]:
        rid = payments.iloc[i]["payment_id"]
        old = payments.iloc[i]["amount_paid"]
        new_val = -abs(old) if pd.notna(old) else -100.0
        payments.at[payments.index[i], "amount_paid"] = new_val
        log.add("payments", rid, "negative_payment_amount",
                "amount_paid is negative, which is not a valid payment amount", "high")

    # G3: total payments on a loan exceed the loan amount by an unrealistic margin
    loan_totals = payments.groupby("loan_id")["amount_paid"].sum()
    candidate_loans = loans[loans["loan_id"].isin(loan_totals.index)]
    if len(candidate_loans) > 0:
        n_pick = max(1, int(round(len(candidate_loans) * rate)))
        picked_loans = candidate_loans.sample(n=min(n_pick, len(candidate_loans)), random_state=rng.integers(0, 2**31 - 1))
        for loan_id, loan_amount in zip(picked_loans["loan_id"], picked_loans["loan_amount"]):
            loan_payment_idx = payments.index[payments["loan_id"] == loan_id]
            if len(loan_payment_idx) == 0:
                continue
            bump_idx = rng.choice(loan_payment_idx)
            inflated = round(loan_amount * rng.uniform(1.5, 3.0), 2)
            payments.at[bump_idx, "amount_paid"] = inflated
            log.add("payments", payments.at[bump_idx, "payment_id"], "overpayment_exceeds_loan",
                    f"amount_paid ({inflated}) combined with other payments unrealistically exceeds loan_amount ({loan_amount})",
                    "medium")

    # G4: defaulted loan has a suspicious payment status (e.g. all "On-time")
    defaulted_loans = loans[loans["loan_status"] == "Defaulted"]
    if len(defaulted_loans) > 0:
        n_pick = max(1, int(round(len(defaulted_loans) * rate)))
        picked = defaulted_loans.sample(n=min(n_pick, len(defaulted_loans)), random_state=rng.integers(0, 2**31 - 1))
        for loan_id in picked["loan_id"]:
            idx = payments.index[payments["loan_id"] == loan_id]
            if len(idx) == 0:
                continue
            payments.loc[idx, "payment_status"] = "On-time"
            log.add("loans", loan_id, "defaulted_loan_all_ontime_payments",
                    "Loan is marked Defaulted but all of its payments are marked On-time", "medium")

    return loans, applications, payments

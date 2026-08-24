"""
generate_data.py
------------------
CLI entry point. Orchestrates the whole pipeline in the required order:

    customers -> applications -> loans -> payments -> credit_history
    -> risk_assessments -> error injection (final pass) -> CSV output

Usage
-----
    python generate_data.py \
        --customers 100000 \
        --applications 150000 \
        --loans 100000 \
        --payments 1000000 \
        --credit-history 300000 \
        --seed 42 \
        --output data/raw

Run with no arguments to generate a small (1k-customer) sample dataset
using the defaults in config.py -- useful for a quick smoke test before
committing to a 100k+ row run.
"""

from __future__ import annotations

import argparse
import os
import time

import numpy as np
import pandas as pd

from data_source.config import Config, DEFAULT_ERROR_CONFIG, EMPLOYMENT_STATUS_VARIANTS, \
    APPLICATION_STATUS_VARIANTS, PAYMENT_STATUS_VARIANTS
from data_source.generators.customers import generate_customers
from data_source.generators.applications import generate_applications
from data_source.generators.loans import generate_loans
from data_source.generators.payments import generate_payments
from data_source.generators.credit_history import generate_credit_history
from data_source.generators.risk_assessments import generate_risk_assessments
from data_source.quality.error_injection import (
    IssueLog,
    inject_missing_values,
    inject_duplicates,
    inject_invalid_numeric,
    inject_impossible_age,
    inject_inconsistent_categories,
    inject_invalid_dates_applications,
    inject_invalid_dates_loans,
    inject_invalid_dates_payments,
    inject_orphan_fk,
    inject_business_rule_violations,
)


def parse_args() -> Config:
    p = argparse.ArgumentParser(description="Generate synthetic credit-risk data for the analytics platform.")
    p.add_argument("--customers", type=int, default=1_000)
    p.add_argument("--applications", type=int, default=1_500)
    p.add_argument("--loans", type=int, default=1_000)
    p.add_argument("--payments", type=int, default=10_000)
    p.add_argument("--credit-history", type=int, default=3_000, dest="credit_history")
    p.add_argument("--seed", type=int, default=42)
    p.add_argument("--output", type=str, default="data/raw")
    p.add_argument("--as-of-date", type=str, default="2026-08-24", dest="as_of_date")
    p.add_argument("--history-start-year", type=int, default=2019, dest="history_start_year")
    # Error rate overrides (optional, default to config.DEFAULT_ERROR_CONFIG)
    for key, default in DEFAULT_ERROR_CONFIG.items():
        p.add_argument(f"--rate-{key.replace('_', '-')}", type=float, default=default, dest=key)
    args = p.parse_args()

    error_config = {key: getattr(args, key) for key in DEFAULT_ERROR_CONFIG}

    return Config(
        n_customers=args.customers,
        n_applications=args.applications,
        n_loans=args.loans,
        n_payments=args.payments,
        n_credit_history=args.credit_history,
        seed=args.seed,
        error_config=error_config,
        output_dir=args.output,
        as_of_date=args.as_of_date,
        history_start_year=args.history_start_year,
    )


def run(cfg: Config) -> dict:
    t0 = time.time()
    rng = np.random.default_rng(cfg.seed)
    log = IssueLog(injected_at=cfg.as_of_date)

    print(f"[1/7] Generating {cfg.n_customers:,} customers...")
    customers = generate_customers(rng, cfg.n_customers, cfg.history_start_year, cfg.as_of_date)

    print(f"[2/7] Generating {cfg.n_applications:,} applications...")
    applications = generate_applications(rng, cfg.n_applications, customers, cfg.as_of_date)

    print(f"[3/7] Generating up to {cfg.n_loans:,} loans (from approved applications)...")
    loans = generate_loans(rng, cfg.n_loans, applications, customers, cfg.as_of_date)
    n_approved = (applications["application_status"] == "Approved").sum()
    if len(loans) < cfg.n_loans:
        print(f"      note: only {n_approved:,} applications were Approved, "
              f"so {len(loans):,} loans were generated (capped by approved-application supply).")

    print(f"[4/7] Generating up to {cfg.n_payments:,} payments...")
    payments = generate_payments(rng, cfg.n_payments, loans, cfg.as_of_date)

    print(f"[5/7] Generating ~{cfg.n_credit_history:,} credit history records...")
    credit_history = generate_credit_history(rng, cfg.n_credit_history, customers, cfg.history_start_year, cfg.as_of_date)

    print("[6/7] Generating risk assessments...")
    risk_assessments = generate_risk_assessments(
        rng, customers, loans, applications, credit_history, cfg.as_of_date, cfg.model_version
    )

    print("[7/7] Injecting data quality issues...")
    ec = cfg.error_config

    # --- A. Missing values (per entity, on realistic columns) ---
    customers = inject_missing_values(customers, "customers",
        ["annual_income", "employment_status", "monthly_income", "city"], ec["missing_values"], rng, log)
    applications = inject_missing_values(applications, "applications",
        ["requested_amount", "loan_purpose"], ec["missing_values"], rng, log)
    loans = inject_missing_values(loans, "loans",
        ["loan_amount", "interest_rate"], ec["missing_values"], rng, log)
    payments = inject_missing_values(payments, "payments",
        ["payment_date", "amount_paid"], ec["missing_values"], rng, log)
    credit_history = inject_missing_values(credit_history, "credit_history",
        ["credit_score", "total_outstanding_debt"], ec["missing_values"], rng, log)

    # --- B. Duplicates ---
    customers = inject_duplicates(customers, "customers", ec["duplicates"], rng, log)
    applications = inject_duplicates(applications, "applications", ec["duplicates"], rng, log)
    payments = inject_duplicates(payments, "payments", ec["duplicates"], rng, log)

    # --- C. Invalid numerical values ---
    customers = inject_invalid_numeric(customers, "customers", "annual_income", ec["invalid_values"], rng, log, mode="negative")
    customers = inject_impossible_age(customers, ec["invalid_values"], rng, log)
    loans = inject_invalid_numeric(loans, "loans", "loan_amount", ec["invalid_values"], rng, log, mode="negative")
    loans = inject_invalid_numeric(loans, "loans", "interest_rate", ec["invalid_values"], rng, log, mode="negative")
    payments = inject_invalid_numeric(payments, "payments", "amount_paid", ec["invalid_values"], rng, log, mode="negative")
    credit_history = inject_invalid_numeric(credit_history, "credit_history", "credit_score",
                                             ec["invalid_values"], rng, log, mode="out_of_range", bound=850)

    # --- D. Inconsistent categorical values ---
    customers = inject_inconsistent_categories(customers, "customers", "employment_status",
                                                EMPLOYMENT_STATUS_VARIANTS, ec["inconsistent_categories"], rng, log)
    applications = inject_inconsistent_categories(applications, "applications", "application_status",
                                                   APPLICATION_STATUS_VARIANTS, ec["inconsistent_categories"], rng, log)
    payments = inject_inconsistent_categories(payments, "payments", "payment_status",
                                               PAYMENT_STATUS_VARIANTS, ec["inconsistent_categories"], rng, log)

    # --- E. Invalid dates ---
    applications = inject_invalid_dates_applications(applications, ec["invalid_dates"], rng, log)
    loans = inject_invalid_dates_loans(loans, ec["invalid_dates"], rng, log)
    payments = inject_invalid_dates_payments(payments, loans, ec["invalid_dates"], rng, log)

    # --- F. Orphan foreign keys ---
    applications = inject_orphan_fk(applications, "applications", "customer_id", "CUST", ec["orphan_records"], rng, log)
    loans = inject_orphan_fk(loans, "loans", "application_id", "APP", ec["orphan_records"], rng, log)
    payments = inject_orphan_fk(payments, "payments", "loan_id", "LOAN", ec["orphan_records"], rng, log)
    credit_history = inject_orphan_fk(credit_history, "credit_history", "customer_id", "CUST", ec["orphan_records"], rng, log)

    # --- G. Business rule violations ---
    loans, applications, payments = inject_business_rule_violations(
        loans, applications, payments, ec["business_rule_violations"], rng, log
    )

    os.makedirs(cfg.output_dir, exist_ok=True)
    outputs = {
        "customers.csv": customers,
        "applications.csv": applications,
        "loans.csv": loans,
        "payments.csv": payments,
        "credit_history.csv": credit_history,
        "risk_assessments.csv": risk_assessments,
        "data_quality_issues.csv": log.to_dataframe(),
    }
    for filename, df in outputs.items():
        path = os.path.join(cfg.output_dir, filename)
        df.to_csv(path, index=False)

    elapsed = time.time() - t0
    print(f"\nDone in {elapsed:.1f}s. Files written to '{cfg.output_dir}/':")
    for filename, df in outputs.items():
        print(f"  {filename:<28} {len(df):>10,} rows")

    return outputs


if __name__ == "__main__":
    cfg = parse_args()
    run(cfg)

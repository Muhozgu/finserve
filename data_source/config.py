"""
config.py
---------
Central configuration for the Credit Risk Analytics & Data Quality Platform
synthetic data generator.

All tunable parameters live here so the rest of the codebase never hardcodes
magic numbers. `Config` is a plain dataclass that gets built from CLI args
in generate_data.py, but can also be constructed directly for notebook /
unit-test use.
"""

from dataclasses import dataclass, field
from typing import Dict


# ---------------------------------------------------------------------------
# Error injection configuration
# ---------------------------------------------------------------------------
# Each key is an error *category*. The value is the fraction of eligible
# records (per relevant entity) that will receive that category of error.
# These are intentionally kept low (1-5%) so the dataset stays "mostly
# valid" -- that's what makes it realistic and what gives dbt tests
# something meaningful (but not overwhelming) to catch.
DEFAULT_ERROR_CONFIG: Dict[str, float] = {
    "missing_values": 0.02,
    "duplicates": 0.01,
    "invalid_values": 0.01,
    "inconsistent_categories": 0.02,
    "invalid_dates": 0.005,
    "orphan_records": 0.005,
    "business_rule_violations": 0.005,
}

# Categorical "messiness" maps: canonical value -> list of dirty variants
# that a real operational system would plausibly produce (typos, casing,
# abbreviations, legacy codes, etc). Used by quality/error_injection.py.
EMPLOYMENT_STATUS_VARIANTS = {
    "Employed": ["employed", "EMPLOYED", "Emp.", "Full-time", "full time", "FT"],
    "Self-employed": ["self employed", "SELF-EMPLOYED", "Self Employed", "Freelance"],
    "Unemployed": ["unemployed", "UNEMPLOYED", "Not Employed", "N/A"],
    "Retired": ["retired", "RETIRED", "Ret."],
    "Student": ["student", "STUDENT", "Stu."],
}

APPLICATION_STATUS_VARIANTS = {
    "Approved": ["approved", "APPROVED", "Appr.", "APPR"],
    "Rejected": ["rejected", "REJECTED", "Declined", "declined", "DECL"],
    "Pending": ["pending", "PENDING", "In Review", "in review"],
    "Withdrawn": ["withdrawn", "WITHDRAWN", "Cancelled", "cancelled"],
}

PAYMENT_STATUS_VARIANTS = {
    "On-time": ["on time", "ON-TIME", "On Time", "OK"],
    "Late": ["late", "LATE", "Delayed", "delayed"],
    "Missed": ["missed", "MISSED", "NSF", "Failed"],
    "Partial": ["partial", "PARTIAL", "Part Payment"],
}


@dataclass
class Config:
    # ---- volumes ----
    n_customers: int = 1_000
    n_applications: int = 1_500
    n_loans: int = 1_000
    n_payments: int = 10_000
    n_credit_history: int = 3_000

    # ---- reproducibility ----
    seed: int = 42

    # ---- error injection ----
    error_config: Dict[str, float] = field(default_factory=lambda: dict(DEFAULT_ERROR_CONFIG))

    # ---- output ----
    output_dir: str = "data/raw"

    # ---- date bounds for the simulated operational window ----
    history_start_year: int = 2019
    as_of_date: str = "2026-08-24"  # "today" for the simulated system

    # ---- model version tag written to risk_assessments ----
    model_version: str = "synthetic-pd-v1.0"

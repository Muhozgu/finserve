"""
generators/customers.py
------------------------
Generates the CUSTOMERS entity -- the root of the FK graph. Every other
entity ultimately traces back to a customer_id generated here.

Financial realism baked in:
- annual_income is drawn from a lognormal distribution (income is right-skewed
  in the real world: many people cluster around the median, a smaller number
  earn much more).
- employment_status influences the income distribution's center (e.g.
  "Unemployed" customers have near-zero or very low income; "Self-employed"
  has higher variance than "Employed").
- employment_length_years is correlated with age (you can't have 30 years of
  employment history at age 22).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from data_source.utils.helpers import make_ids, random_dates


COUNTRIES = ["Poland", "Germany", "France", "United Kingdom", "Spain", "Italy", "Netherlands"]
COUNTRY_WEIGHTS = [0.35, 0.15, 0.12, 0.12, 0.10, 0.08, 0.08]

CITIES_BY_COUNTRY = {
    "Poland": ["Warsaw", "Krakow", "Wroclaw", "Poznan", "Gdansk"],
    "Germany": ["Berlin", "Munich", "Hamburg", "Frankfurt"],
    "France": ["Paris", "Lyon", "Marseille"],
    "United Kingdom": ["London", "Manchester", "Birmingham"],
    "Spain": ["Madrid", "Barcelona", "Valencia"],
    "Italy": ["Rome", "Milan", "Turin"],
    "Netherlands": ["Amsterdam", "Rotterdam", "Utrecht"],
}

EMPLOYMENT_STATUSES = ["Employed", "Self-employed", "Unemployed", "Retired", "Student"]
EMPLOYMENT_WEIGHTS = [0.62, 0.14, 0.08, 0.12, 0.04]

# Median annual income (EUR-equivalent) and lognormal sigma per employment status.
INCOME_PARAMS = {
    "Employed":       {"median": 48_000, "sigma": 0.42},
    "Self-employed":  {"median": 45_000, "sigma": 0.65},
    "Unemployed":     {"median": 6_000,  "sigma": 0.55},
    "Retired":        {"median": 24_000, "sigma": 0.35},
    "Student":        {"median": 9_000,  "sigma": 0.50},
}


def generate_customers(rng: np.random.Generator, n: int, history_start_year: int, as_of_date: str) -> pd.DataFrame:
    customer_id = make_ids("CUST", n)

    # --- demographics ---
    # Ages 18-85, roughly bell-shaped around working-age adults.
    age = np.clip(rng.normal(loc=42, scale=14, size=n), 18, 85).round().astype(int)
    as_of = pd.Timestamp(as_of_date)
    date_of_birth = as_of - pd.to_timedelta(age * 365.25, unit="D")
    date_of_birth = date_of_birth.normalize()

    gender = rng.choice(["Female", "Male", "Non-binary"], size=n, p=[0.49, 0.49, 0.02])

    country = rng.choice(COUNTRIES, size=n, p=COUNTRY_WEIGHTS)
    city = np.array([rng.choice(CITIES_BY_COUNTRY[c]) for c in country])

    first_names = ["Anna", "Jan", "Maria", "Piotr", "Sofia", "Marco", "Emma", "Lucas",
                   "Julia", "Tomasz", "Laura", "Hugo", "Nina", "Diego", "Elena", "Felix"]
    last_names = ["Kowalski", "Nowak", "Muller", "Garcia", "Rossi", "Dubois", "Smith",
                  "Jansen", "Silva", "Andersson", "Novak", "Fischer", "Moreau", "Lopez"]
    first_name = rng.choice(first_names, size=n)
    last_name = rng.choice(last_names, size=n)

    # --- employment & income ---
    employment_status = rng.choice(EMPLOYMENT_STATUSES, size=n, p=EMPLOYMENT_WEIGHTS)

    # employment_length capped by (age - 18), can't have worked longer than adult life
    max_possible_length = np.clip(age - 18, 0, None)
    raw_length = rng.gamma(shape=2.0, scale=4.0, size=n)
    employment_length_years = np.minimum(raw_length, max_possible_length).round(1)
    # Unemployed / Student customers realistically have low/zero current employment length
    employment_length_years = np.where(
        np.isin(employment_status, ["Unemployed", "Student"]),
        np.round(rng.uniform(0, 1.5, size=n), 1),
        employment_length_years,
    )

    annual_income = np.zeros(n)
    for status, params in INCOME_PARAMS.items():
        mask = employment_status == status
        count = mask.sum()
        if count == 0:
            continue
        mu = np.log(params["median"])
        sample = rng.lognormal(mean=mu, sigma=params["sigma"], size=count)
        annual_income[mask] = sample

    # Slight income bump for longer employment tenure (loyalty/seniority effect)
    annual_income = annual_income * (1 + 0.01 * np.minimum(employment_length_years, 20))
    annual_income = np.round(annual_income, 2)
    monthly_income = np.round(annual_income / 12, 2)

    created_at = random_dates(rng, f"{history_start_year}-01-01", as_of_date, n)

    df = pd.DataFrame({
        "customer_id": customer_id,
        "first_name": first_name,
        "last_name": last_name,
        "date_of_birth": date_of_birth.strftime("%Y-%m-%d"),
        "gender": gender,
        "country": country,
        "city": city,
        "employment_status": employment_status,
        "employment_length_years": employment_length_years,
        "annual_income": annual_income,
        "monthly_income": monthly_income,
        "created_at": created_at.strftime("%Y-%m-%d"),
    })
    return df

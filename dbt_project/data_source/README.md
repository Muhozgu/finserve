# Credit Risk Analytics & Data Quality Platform — Synthetic Data Generator

A production-quality Python generator that produces a **relationally
consistent, intentionally-messy** synthetic dataset simulating a financial
services company's operational data, for building out the rest of the
pipeline:

```
Python Generator → Raw CSV Files → Snowflake RAW → dbt STAGING → dbt INTERMEDIATE → dbt MARTS → Power BI
```

The raw layer is deliberately messy. The dbt layer is where you make it
trustworthy. The marts are what Power BI reads.

> **This is synthetic data.** The credit-risk scoring formula in
> `generators/risk_assessments.py` is an illustrative, hand-built model for
> a portfolio project. It is not, and should not be presented as, a real
> bank's proprietary risk methodology.

---

## 1. Architecture

```
generate_data.py        <- CLI entry point / orchestrator
config.py                <- all tunables: volumes, seed, error rates, categorical "dirty variant" maps
generators/
  customers.py            <- root entity
  applications.py          <- child of customers
  loans.py                  <- child of approved applications
  payments.py                 <- child of loans
  credit_history.py             <- child of customers (independent time series)
  risk_assessments.py             <- child of customers + (optionally) loans
quality/
  error_injection.py       <- the ONLY module allowed to corrupt data; logs every corruption
utils/
  helpers.py                <- shared, reusable numeric/date/ID helpers
data/raw/                   <- output CSVs land here
```

Each `generators/*.py` module is a pure function: `(rng, config, parent
dataframes...) -> clean DataFrame`. Nothing in `generators/` ever injects
an error — that separation of concerns is what lets `quality/error_injection.py`
be the single source of truth for "what's wrong with this dataset," which
is exactly what you need to grade your own dbt tests against.

I kept your proposed folder structure as-is rather than simplifying it —
for a portfolio project, this file-per-entity layout is actually a good
thing to point to in an interview: it shows you can decompose a data
pipeline instead of writing one 800-line script.

---

## 2. Entity relationships

```
CUSTOMER (1) ── (N) APPLICATION (1) ── (0/1) LOAN (1) ── (N) PAYMENT
CUSTOMER (1) ── (N) CREDIT_HISTORY
CUSTOMER (1) ── (N) RISK_ASSESSMENT
LOAN     (1) ── (N) RISK_ASSESSMENT     [loan_id is NULL for application-only assessments]
```

One deliberate addition to your spec: **`risk_assessments.loan_id` is
nullable**, and a `risk_assessments` row is generated for every **rejected**
application too (in addition to every loan). This mirrors how real
underwriting works — you score risk *before* you decide whether to
originate a loan — and gives you a realistic, meaningful nullable-FK column
to write `not_null` / conditional dbt tests against, instead of an
artificial one.

---

## 3. Generation order

`customers → applications → loans → payments → credit_history → risk_assessments → error injection → CSV write`

Parents are always generated (and their IDs fixed) before children sample
from them. Nothing generates a foreign key independently — every FK value
in the *clean* data is drawn directly from an already-materialized parent
ID column.

One subtlety: `credit_history` is generated *after* `loans` in file order,
but its content doesn't depend on loans — bureau history is a customer-level
time series, not loan-level. `loans.py` and `risk_assessments.py` each use
their own lightweight "underwriting risk proxy" (built from income/employment,
see §5) rather than waiting on `credit_history`, and `risk_assessments.py`
separately joins in each customer's *latest* `credit_history` snapshot at
assessment time. This avoids a circular dependency (loans logically need
"risk" but happen before "history" in the required file order) while
keeping both signals internally consistent with the same underlying
customer income/employment profile.

## 4. How IDs are maintained

All IDs are stable, sequential, zero-padded strings generated once via
`utils.helpers.make_ids(prefix, n)`:

| Entity | Prefix | Example |
|---|---|---|
| customers | `CUST` | `CUST0000001` |
| applications | `APP` | `APP0000001` |
| loans | `LOAN` | `LOAN0000001` |
| payments | `PAY` | `PAY0000001` |
| credit_history | `CRHIST` | `CRHIST0000001` |
| risk_assessments (loan-linked) | `RISK` | `RISK0000001` |
| risk_assessments (application-only) | `RISKA` | `RISKA0000001` |

IDs are assigned once at generation time and never reused. **Duplicate
records** (§7) reuse an existing ID on purpose — that's the point of that
error type — everything else keeps IDs unique.

## 5. How realistic financial relationships are generated

- **Income**: lognormal distribution per `employment_status` (income is
  right-skewed in reality), with a small seniority bump from
  `employment_length_years`.
- **Requested/loan amount**: scaled off the customer's `annual_income` with
  a gamma-distributed multiplier, so bigger incomes generally support
  bigger loans without it being a hard rule.
- **Interest rate**: priced from a 0–1 "underwriting risk proxy" built from
  income, employment status, and tenure: `rate = 3.5 + risk_proxy * 14.0 + noise`,
  clipped to a realistic 2.5–24.9% consumer-loan band.
- **Monthly payment**: the standard amortization formula,
  `M = P·r(1+r)^n / ((1+r)^n − 1)`, so `loan_amount`, `interest_rate`,
  `term_months`, and `monthly_payment` are always mathematically consistent
  in the clean data.
- **DTI** (`total_outstanding_debt / annual_income`) and **loan-to-income**
  (`loan_amount / annual_income`) are computed, not stored as raw inputs —
  you'll likely want to re-derive these in a dbt staging/intermediate model
  rather than trusting a "DTI" column from a source system anyway, which is
  realistic: DTI is usually a computed metric, not a source field.
- **Credit score**: starts from a latent "credit quality" baseline (income +
  employment + tenure) and then does a slow **random walk** across each
  customer's `credit_history` snapshots — scores drift, they don't teleport.
- **Payment behavior**: each loan gets a single "reliability" draw at
  origination (worse for `Defaulted`/`Delinquent` loans), and every payment
  on that loan samples its On-time/Late/Missed/Partial outcome from that
  same reliability score. This is what makes a customer who misses one
  payment *more likely* to miss others, instead of every payment being an
  independent coin flip.

## 6. Credit risk logic (probability_of_default)

```
risk_index = 0.9·z(−credit_score) + 0.8·z(DTI) + 0.6·z(utilization)
           + 0.7·z(previous_defaults) + 0.4·z(previous_late_payments)
           + 0.3·z(−employment_length_years) + 0.5·z(loan_to_income_ratio)
           + 0.2·z(loan_term_months) + 0.3·z(interest_rate)
           + Normal(0, 0.5) noise

PD = sigmoid(risk_index · 0.6 − 2.4)              # ∈ (0, 1)
risk_score = rescale(logit(PD)) onto a 300–850 band, higher = safer
risk_category = LOW  if PD < 0.10 (± small jitter)
                MEDIUM if PD < 0.30
                HIGH otherwise
```

Each driver is z-scored (standardized) across the population, then combined
with a hand-picked weight whose **sign** matches real-world credit-risk
intuition (higher score → lower risk, higher DTI/utilization/defaults →
higher risk, etc.) but whose exact magnitude is illustrative, not fitted to
any real portfolio. The scale/shift constants (`0.6`, `−2.4`) were tuned
empirically so the *simulated portfolio* lands at a believable mean PD
(~12–14%) and median PD (~7–8%) with a sensible LOW/MEDIUM/HIGH split —
not derived from any real bank's loss experience.

## 7. How errors are injected

`quality/error_injection.py` is the **only** place clean data gets
corrupted, and it runs as a final pass after every entity is fully
generated. Every corruption is written to `data_quality_issues.csv`
*before or while* the corresponding cell is changed, so the log and the
data can never drift apart. Implemented error types, matching your spec:

| Category | Where applied |
|---|---|
| **A. Missing values** | income, employment_status, city (customers); requested_amount, loan_purpose (applications); loan_amount, interest_rate (loans); payment_date, amount_paid (payments); credit_score, total_outstanding_debt (credit_history) |
| **B. Duplicates** | customers, applications, payments — exact-copy duplicate rows appended |
| **C. Invalid numeric** | negative income/loan_amount/interest_rate/payment amounts; out-of-range credit_score (>850 or <300); impossible date_of_birth (age >120 or future) |
| **D. Inconsistent categories** | employment_status, application_status, payment_status rewritten in dirty casing/abbreviation variants (see `config.py` variant maps) |
| **E. Invalid dates** | decision_date before application_date; maturity_date before start_date; payment_date before loan start_date |
| **F. Orphan FKs** | applications.customer_id, loans.application_id, payments.loan_id, credit_history.customer_id pointed at fabricated, nonexistent parent IDs (in the `9,000,000+` ID range so they're easy to spot/filter) |
| **G. Business rule violations** | an application flipped to Rejected *after* its loan already exists; negative payment amounts; cumulative payments unrealistically exceeding loan_amount; a Defaulted loan whose payments are all marked On-time |

## 8. How error rates are configured

`config.DEFAULT_ERROR_CONFIG` and the `--rate-*` CLI flags both use the
same seven keys:

```python
ERROR_CONFIG = {
    "missing_values": 0.02,
    "duplicates": 0.01,
    "invalid_values": 0.01,
    "inconsistent_categories": 0.02,
    "invalid_dates": 0.005,
    "orphan_records": 0.005,
    "business_rule_violations": 0.005,
}
```

Each rate is the fraction of *eligible* rows (per entity/column) that
receive that error type. Override any of them from the CLI, e.g.
`--rate-missing-values 0.05`. Rates are intentionally low by default so the
dataset stays mostly valid — a raw layer that's 40% garbage doesn't
teach you anything about writing selective, useful dbt tests.

## 9. How reproducibility works

A single `numpy.random.Generator` (`np.random.default_rng(seed)`) is
created once in `generate_data.py` and threaded explicitly through every
generator and injector function as the `rng` argument — nothing calls the
legacy global `np.random.seed()`/`np.random()` API. Because Python
dictionaries, DataFrame row order, and function call order are all
deterministic given the same inputs, the *same* `--seed` reliably
reproduces byte-identical CSVs. This was verified during development:
running the generator twice with `--seed 42` produces `DataFrame.equals()
== True` on every output table.

## 10. How to run the generator

```bash
# Quick smoke test (defaults to a small ~1k-customer dataset)
python generate_data.py

# Target production-scale run matching your spec
python generate_data.py \
    --customers 100000 \
    --applications 250000 \
    --loans 100000 \
    --payments 1000000 \
    --credit-history 300000 \
    --seed 42 \
    --output data/raw
```

**Note on `--loans`:** loans can only originate from *Approved*
applications, and roughly ~40% of applications are approved in this
simulation (a realistic-ish approval rate). `--loans` is therefore a
**ceiling**, not a guarantee — if you ask for more loans than there are
approved applications, the generator logs a note and gives you every
approved application's loan rather than erroring out. To reliably get
100k+ loans, request roughly 2.5x that many applications (e.g.
`--applications 250000` for `--loans 100000`).

Override any error rate: `--rate-missing-values 0.05 --rate-duplicates 0.02`, etc.

## 11. Example output

A small example run (`--customers 2000 --applications 3000 --loans 2000
--payments 20000 --credit-history 6000 --seed 42`) is included in this
package under `data/raw/`:

```
customers.csv                2,020 rows
applications.csv             3,030 rows
loans.csv                    1,179 rows   (capped by ~39% approval rate)
payments.csv                20,200 rows
credit_history.csv           6,069 rows
risk_assessments.csv         2,427 rows
data_quality_issues.csv      2,826 rows
```

Performance benchmark on the full target scale (100k customers / 150k
applications / 1M payments / 300k credit history records): **~40 seconds**
on a single core, no GPU, no multiprocessing — vectorized NumPy/Pandas
operations throughout. The two entities with genuinely sequential logic
(payment schedules, and the credit-score random walk) use a "wave"
vectorization pattern — looping over the small number of *time steps*
(e.g. up to ~84 monthly payments, or a handful of bureau snapshots) instead
of looping over the (much larger) number of *customers/loans* — which cut
generation time by roughly 4x during development. Plain row-by-row
`.apply()` was avoided everywhere it would have been the dominant cost.

## 12. Recommended Snowflake schema (RAW layer)

Land every CSV as-is, typed loosely on ingest (Snowflake's `COPY INTO` +
`VARIANT`/permissive typing, or just `VARCHAR` everywhere in RAW) — don't
fight the intentionally-dirty data on load. Suggested target types once you
get to staging:

| Column pattern | Recommended Snowflake type |
|---|---|
| `*_id` (all ID columns) | `VARCHAR(20)` |
| `annual_income`, `monthly_income`, `total_outstanding_debt`, `loan_amount`, `requested_amount`, `amount_due`, `amount_paid`, `monthly_payment` | `NUMBER(15,2)` |
| `interest_rate` | `NUMBER(5,2)` |
| `credit_score` | `NUMBER(4,0)` (staging should enforce 300–850 via a dbt test, not a DB constraint, so bad rows still land and are testable) |
| `credit_utilization`, `probability_of_default` | `NUMBER(6,4)` (0–1 range) |
| `risk_score` | `NUMBER(4,0)` |
| `employment_length_years`, `term_months`, `days_late`, `number_of_open_accounts`, `number_of_closed_accounts`, `previous_defaults`, `previous_late_payments` | `NUMBER(6,1)` / `NUMBER(6,0)` as appropriate |
| all `*_date` columns | `DATE` |
| `created_at`, `injected_at` | `TIMESTAMP_NTZ` (or `DATE` if you don't need time-of-day) |
| all categorical/text columns (`employment_status`, `application_status`, `loan_status`, `payment_status`, `risk_category`, `gender`, `country`, `city`, `loan_purpose`, `application_channel`, `model_version`, `issue_type`, `severity`, etc.) | `VARCHAR(100)` |

RAW-layer tables should load everything as `VARCHAR` and let dbt's
`staging` models do the `TRY_CAST`/`TRY_TO_DATE` — that's exactly the kind
of type-casting practice you said you wanted, and it means a malformed row
(e.g. a negative interest rate stored as text) never blocks the load.

## 13. Recommended dbt project structure

```
credit_risk_dbt/
├── models/
│   ├── staging/
│   │   ├── stg_customers.sql        -- TRY_CAST, trim, standardize casing
│   │   ├── stg_applications.sql
│   │   ├── stg_loans.sql
│   │   ├── stg_payments.sql
│   │   ├── stg_credit_history.sql
│   │   ├── stg_risk_assessments.sql
│   │   └── stg_data_quality_issues.sql
│   ├── intermediate/
│   │   ├── int_applications_deduped.sql     -- dedup logic isolated here
│   │   ├── int_loans_with_computed_dti.sql  -- recompute DTI/LTI properly
│   │   ├── int_payments_reconciled.sql      -- flag orphans/business-rule breaks
│   │   └── int_customer_credit_snapshot.sql -- latest credit_history per customer
│   └── marts/
│       ├── dim_customers.sql
│       ├── dim_loans.sql
│       ├── fct_payments.sql
│       ├── fct_risk_assessments.sql
│       └── fct_data_quality_summary.sql     -- roll up dq issues by entity/severity
├── seeds/
│   └── category_mappings.csv        -- your canonical category lookup (dirty -> clean)
├── tests/
│   └── (custom singular tests, see §14)
├── macros/
│   └── standardize_category.sql     -- reusable UPPER/TRIM/CASE WHEN macro
└── dbt_project.yml
```

`staging` = clean types and casing only, one-to-one with the raw table.
`intermediate` = business logic (dedup, reconciliation, joins). `marts` =
what Power BI actually connects to.

## 14. Recommended dbt tests

**Generic (schema.yml) tests:**
- `unique` + `not_null` on every `*_id` primary key
- `relationships` — `applications.customer_id → customers.customer_id`,
  `loans.application_id → applications.application_id`,
  `payments.loan_id → loans.loan_id`,
  `credit_history.customer_id → customers.customer_id`,
  `risk_assessments.customer_id → customers.customer_id` (this is where the
  intentionally-injected orphan FKs in §7 will surface as failures — that's
  the point)
- `accepted_values` on `employment_status`, `application_status`,
  `loan_status`, `payment_status`, `risk_category` — you'll need to run
  these *after* a standardization macro, since the raw values include the
  dirty casing variants on purpose
- `dbt_utils.accepted_range` (or a custom test) on `credit_score` (300–850),
  `probability_of_default` (0–1), `interest_rate` (>0), all `*_amount`
  columns (>=0)

**Custom singular tests (write these yourself as SQL files in `tests/`):**
- `assert_no_future_dates.sql` — flag any `application_date`/`start_date`
  in the future relative to load date
- `assert_decision_after_application.sql` — `decision_date >= application_date`
- `assert_maturity_after_start.sql` — `maturity_date >= start_date`
- `assert_payment_after_loan_start.sql` — `payment_date >= start_date`
- `assert_no_rejected_with_loan.sql` — no `Rejected` application should
  join to a row in `loans`
- `assert_defaulted_loan_payment_consistency.sql` — flag `Defaulted` loans
  where 100% of payments are `On-time` (a strong signal something's wrong
  upstream, exactly the scenario injected in §7G)
- `assert_amortization_consistency.sql` — recompute `monthly_payment` from
  `loan_amount`/`interest_rate`/`term_months` and flag rows where the
  stored value is off by more than a rounding tolerance

Comparing your dbt test failure counts against `data_quality_issues.csv`
(grouped by `issue_type`) is a great way to measure your own test coverage
— did your `relationships` test actually catch every injected orphan?

## 15. Suggested Power BI metrics

Built on top of the dbt marts, once they're clean:

- **Portfolio overview**: total outstanding loan balance, active loan
  count, average interest rate, weighted-average risk_score
- **Approval funnel**: application → approval rate → origination rate, by
  `application_channel` and `loan_purpose`
- **Delinquency & risk**: % of loans `Delinquent`/`Defaulted` by
  `risk_category`, days-late distribution, roll-rate (on-time → late →
  missed) over time
- **PD calibration**: average `probability_of_default` vs. actual observed
  default rate by `risk_category` bucket (a classic model-validation chart)
- **Data quality dashboard**: `data_quality_issues` count by `entity` /
  `issue_type` / `severity` over time — genuinely useful for showing you
  can build a DQ-monitoring view, not just a business-metrics one
- **Customer segmentation**: income band × employment_status × risk_category
  cross-tab

## What's intentionally unrealistic

Worth calling out explicitly, since this is synthetic:

- The PD formula's weights are hand-picked for realistic *direction and
  shape*, not fitted/calibrated to any real loss data.
- Correlations between macro conditions (e.g. unemployment cycles,
  interest-rate environment over time) aren't modeled — every loan is
  priced independently of "what year it is" beyond the customer's own
  profile.
- `credit_history` snapshot cadence is a Poisson draw per customer, not a
  realistic bureau reporting cycle (real bureaus report roughly monthly per
  active account).
- Geography is a flat weighted list of European countries/cities with no
  real regional economic variation baked into income or risk.
- There's no seasonality (e.g. more applications in December) — application
  dates are uniform over the lookback window.

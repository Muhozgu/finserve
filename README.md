# FinServe - Financial Operations & Risk Analytics Platform

An AI-powered process automation platform for financial operations, combining a Python/Pandas ETL pipeline, Snowflake + dbt data warehousing, and Llama 3-driven agents orchestrated via n8n and Power Automate.

> **Note:** FinServe was built as a portfolio/challenge project simulating financial operations workflows (credit memos, ticket triage, portfolio reporting).

---

## Project Goals

- Practice building a full data + AI automation pipeline end to end (ingestion → warehouse → transformation → agents → reporting)
- Automate realistic financial ops tasks: credit memo generation, support ticket triage, portfolio reporting
- Apply a medallion (Bronze/Silver/Gold) architecture with dbt
- Orchestrate LLM-driven agents (Llama 3) into existing operational workflows via n8n and Power Automate
- Produce a Power BI report suitable for a portfolio walkthrough

## Tech Stack

| Layer | Tool |
|---|---|
| Language / ETL | Python (Pandas) |
| Version control | Git / GitHub |
| Data storage | MS Access, PostgreSQL, Snowflake |
| Transformation | dbt (Bronze → Silver → Gold) |
| AI / LLM | Llama 3 (modular agents) |
| Workflow automation | n8n, Power Automate |
| Reporting / modeling | Power BI |
| Containerization | Docker |

## Architecture

```
CSV / XLSX               Python ETL              Snowflake                    Agents / Automation           Power BI
┌─────────────┐  raw    ┌───────────┐  dbt      ┌───────────┐    dbt        ┌─────────────────────┐      ┌───────────┐
│ Financial   │ ───────▶│  BRONZE   │ ─────────▶│  SILVER   │ ─────────────▶│  GOLD (star schema) │      │  Power BI │
│ data sources│         │ (raw load)│           │ (typed,   │               │  + modular agents:   │─────▶│  Reports  │
└─────────────┘         └───────────┘           │  cleaned) │               │  - Credit Memo Agent │      └───────────┘
                                                  └───────────┘               │  - Ticket Triage     │
                                                                              │  - Portfolio Report  │
                                                                              │  (Llama 3, via n8n / │
                                                                              │   Power Automate)    │
                                                                              └─────────────────────┘
```

## Dataset

FinServe ingests financial operations data from CSV/XLSX sources, the kinds of inputs typical in banking/finance back-office workflows:

### Raw inputs
Financial operations records (e.g. credit/portfolio data, support tickets) loaded as-is into the Bronze layer.

### Silver models
Typed and cleaned staging models built with dbt on top of the Snowflake-loaded raw data.

### Gold models
Analytics-ready tables feeding both the Power BI reports and the automation agents (portfolio, credit memo, and ticket data marts).

## Repository Structure

```
finserve/
├── README.md
├── agents/                 # Modular AI agents (credit memo, ticket triage, reporting)
├── data/                    # Python/Pandas ETL pipeline
├── dbt/
│   ├── dbt_project.yml
│   └── models/
│       ├── silver/         # typed/cleaned staging models
│       └── gold/           # star schema, analytics-ready models
├── workflows/               # n8n / Power Automate workflow definitions
├── reports/                 # Power BI report files
└── docker/                  # Docker configuration
```

## Setup

1. **Run the ETL pipeline**
   ```bash
   pip install pandas
   python etl/run_pipeline.py --input-dir data/raw --output-dir data/processed
   ```

2. **Load into Snowflake (Bronze)**
   Load processed CSV/XLSX outputs into a `RAW`/`BRONZE` schema via Snowsight or the `snowflake-connector-python` package.

3. **Run dbt (Silver → Gold)**
   ```bash
   cd dbt
   dbt deps
   dbt run
   dbt test
   ```
   Silver models cast types and standardize columns; Gold models build the analytics-ready schema consumed by both Power BI and the agents.

4. **Deploy the agents**
   ```bash
   docker build -t finserve .
   docker run finserve
   ```
   Import the n8n / Power Automate workflow definitions from `workflows/` to trigger the Credit Memo, Ticket Triage, and Portfolio Reporting agents.

5. **Connect Power BI**
   Open Power BI Desktop → Get Data → Snowflake → point at the `GOLD` schema → Import mode.

## Key Metrics / Outputs Modeled

- Automated credit memo generation
- Support ticket classification and routing
- Portfolio reporting and summary metrics
- Financial operations reporting via Power BI dashboards

## License

MIT License.

# Overview

End-to-end AI automation project built for a fictional financial services company (FinServe) as part of a technical assessment. 
The project identifies three high-impact business problems across the organisation and implements all three as working proof-of-concepts using Python and the Groq API (Llama 3.3 70B).

Problem 1 — AI Credit Memo Generator

Credit analysts at FinServe manually pull data from three internal systems — CRM, core banking, 
and loan applications — into Word templates for every credit committee decision. This solution automates the full pipeline: it extracts and normalises data from all three sources, sends it to an AI model for ratio analysis, risk flagging, and narrative writing, and exports a professionally formatted Word document in seconds.

Problem 2 — Support Ticket Triage

The client support team answers every email individually with no shared knowledge base, 
leading to inconsistent responses and slow handling times. This solution classifies each incoming ticket by category, urgency, and sentiment, flags complex cases for human review, and generates a professional draft response ready for the agent to approve and send.

Problem 3 — Monthly Portfolio Report Generator

Finance and risk teams manually extract data from core banking and CRM systems into Excel each month to produce regulatory and internal reports — a time-consuming process prone to reconciliation errors. This solution automates data extraction, calculates key risk metrics (PAR30, NPL ratio, collection rate), and generates a fully formatted Word report with an AI-written executive summary, risk flags, and recommendations.

Why Problem 1 was chosen as the primary PoC

The credit memo generator sits at the core of FinServe's revenue process, demonstrates the full data-to-document pipeline end-to-end, requires genuine domain knowledge to build well, and produces measurable time savings from day one.
Tech stack: Python · Groq API (Llama 3.3 70B) · python-docx · JSON data simulation · GitHub

# Usage

### Prerequisites

- Python 3.8+
- A free Groq API key from https://console.groq.com

### Installation

pip install groq python-docx

### Configuration

Open `generate_memo.py`, `triage.py` or `generate_report.py` and replace line 12:
GROQ_API_KEY = "your_groq_api_key_here"

Or set it as an environment variable:
* Mac/Linux

export GROQ_API_KEY=your_key_here

* Windows

set GROQ_API_KEY=your_key_here

Problem I
---

### Run
python generate_memo.py

### Output
A Word document is saved in the same folder:
credit_memo_Horizon_Logistics_Sp_z_oo.docx

### Changing the client
Edit `client_data.json` to use a different client.
The script picks up the data automatically on the next run.

### Input data structure
client_data.json contains 3 sections — one per source system:
- source_crm           → client profile and relationship data
- source_core_banking  → financials, repayment history, existing debt
- source_loan_application → loan request, collateral, documents


Problem II
---

### Run

python triage.py

### Output

A structured triage report is printed to the terminal for each ticket:
- Category      (Repayment Enquiry, Complaint, Technical Issue, etc.)
- Urgency       (Low / Medium / High / Critical)
- Sentiment     (Neutral / Frustrated / Angry / Distressed)
- Summary       (one sentence)
- Human review  (flagged yes/no with reason)
- Draft response (ready to copy and send)

A summary table is printed at the end showing totals by category and urgency.

### Adding tickets

Edit `tickets.json` to add or modify incoming support emails.
Each ticket needs: ticket_id, from_name, from_email, company, subject, body.

### Input data structure

tickets.json contains a list of support emails, each with:
- ticket_id    → unique identifier
- from_name    → client name
- from_email   → client email address
- company      → client company name
- subject      → email subject line
- body         → full email content


Problem III
---

### Run

python generate_report.py

### Output

A Word document is saved in the same folder:
monthly_portfolio_report_November_2024.docx

The report includes:
- Executive summary (AI-generated)
- Portfolio overview with month-on-month comparisons
- Breakdown by product and sector
- Credit quality metrics — PAR30, NPL ratio, collection rate
- Overdue bucket analysis
- Top 5 exposures by client
- Risk flags and recommendations (AI-generated)
- Forward-looking outlook (AI-generated)

### Updating the data

Edit `portfolio_data.json` to change the reporting period or figures.
Update the report_period.month and report_period.year fields
to reflect the correct month.

### Input data structure

portfolio_data.json contains 2 sections — one per source system:
- source_core_banking  → active loans, disbursements, overdue buckets, top exposures
- source_crm           → active clients, sector breakdown, pipeline, flagged accounts


# Diagram - Architecture

You can find the fiagram on the link added on the repository description section or inside the attached `Guy_Muhoza_task.pdf/pptx` files on the page number 7.

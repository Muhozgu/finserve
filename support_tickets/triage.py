




import json
import os
from datetime import datetime
from groq import Groq

# CONFIG
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "gsk_xxxxx")  # Replace with your actual Groq API key or set as environment variable
GROQ_MODEL   = "llama-3.3-70b-versatile"
TICKETS_FILE = "tickets.json"

# 1. Load tickets
def load_tickets(filepath: str) -> list:
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)

# 2. Prompts
SYSTEM_PROMPT = """You are an AI assistant for FinServe, a financial services company
specialising in SME and retail lending. You help the client support team by analysing
incoming support tickets and drafting professional responses.

Always respond with ONLY a valid JSON object — no markdown, no code blocks, no extra text."""

def build_prompt(ticket: dict) -> str:
    return f"""
Analyse the following support ticket and respond with a JSON object with exactly these keys:

{{
  "category": "one of: Repayment Enquiry | Complaint | New Loan Enquiry | Technical Issue | Loan Restructuring | Other",
  "urgency": "one of: Low | Medium | High | Critical",
  "sentiment": "one of: Positive | Neutral | Frustrated | Angry | Distressed",
  "summary": "one sentence summarising the client's issue",
  "requires_human": true or false,
  "requires_human_reason": "brief reason if requires_human is true, otherwise empty string",
  "draft_response": "a professional, empathetic draft response from the FinServe support team — address the client by first name, acknowledge their issue, provide helpful next steps, sign off as FinServe Client Support Team"
}}

TICKET:
Ticket ID:  {ticket['ticket_id']}
Received:   {ticket['received_at']}
From:       {ticket['from_name']} <{ticket['from_email']}>
Company:    {ticket['company']}
Subject:    {ticket['subject']}
Message:
{ticket['body']}
"""

# 3. Call Groq API
def call_groq(prompt: str) -> dict:
    client = Groq(api_key=GROQ_API_KEY)
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": prompt}
        ],
        temperature=0.2
    )
    raw = response.choices[0].message.content.strip()
    return json.loads(raw)

# 4. Terminal output helpers
URGENCY_ICONS = {
    "Low":      "🟢",
    "Medium":   "🟡",
    "High":     "🟠",
    "Critical": "🔴"
}

SENTIMENT_ICONS = {
    "Positive":  "😊",
    "Neutral":   "😐",
    "Frustrated":"😤",
    "Angry":     "😠",
    "Distressed":"😟"
}

def print_divider(char="=", width=68):
    print(char * width)

def print_ticket_result(ticket: dict, result: dict, index: int, total: int):
    urgency   = result.get("urgency", "Unknown")
    sentiment = result.get("sentiment", "Unknown")
    requires  = result.get("requires_human", False)

    print_divider("=")
    print(f"  TICKET {index}/{total}  |  {ticket['ticket_id']}  |  {ticket['received_at'][:10]}")
    print_divider("-")
    print(f"  From:    {ticket['from_name']} <{ticket['from_email']}>")
    print(f"  Company: {ticket['company']}")
    print(f"  Subject: {ticket['subject']}")
    print_divider("-")
    print(f"  Category:  {result.get('category', 'Unknown')}")
    print(f"  Urgency:   {URGENCY_ICONS.get(urgency, '')} {urgency}")
    print(f"  Sentiment: {SENTIMENT_ICONS.get(sentiment, '')} {sentiment}")
    print(f"  Summary:   {result.get('summary', '')}")
    print_divider("-")

    if requires:
        print(f"  ⚠️  REQUIRES HUMAN REVIEW")
        print(f"  Reason: {result.get('requires_human_reason', '')}")
        print_divider("-")

    print("  DRAFT RESPONSE:")
    print()
    # Word-wrap the draft response at ~64 chars
    draft = result.get("draft_response", "")
    words = draft.split()
    line  = "  "
    for word in words:
        if len(line) + len(word) + 1 > 66:
            print(line)
            line = "  " + word + " "
        else:
            line += word + " "
    if line.strip():
        print(line)
    print()

# 5. Summary report
def print_summary(tickets: list, results: list):
    print_divider("=")
    print("  TRIAGE SUMMARY")
    print_divider("=")
    print(f"  Total tickets processed: {len(tickets)}")
    print()

    categories = {}
    urgencies  = {}
    human_required = 0

    for r in results:
        cat = r.get("category", "Unknown")
        urg = r.get("urgency",  "Unknown")
        categories[cat] = categories.get(cat, 0) + 1
        urgencies[urg]  = urgencies.get(urg, 0) + 1
        if r.get("requires_human"):
            human_required += 1

    print("  By category:")
    for cat, count in sorted(categories.items()):
        print(f"    {cat:<30} {count}")

    print()
    print("  By urgency:")
    for urg in ["Critical", "High", "Medium", "Low"]:
        count = urgencies.get(urg, 0)
        icon  = URGENCY_ICONS.get(urg, "")
        print(f"    {icon} {urg:<28} {count}")

    print()
    print(f"  Requires human review:         {human_required}/{len(tickets)}")
    print(f"  Can be handled by AI draft:    {len(tickets) - human_required}/{len(tickets)}")
    print_divider("=")

# 6. Main
def main():
    print()
    print_divider("=")
    print("  FinServe AI Support Ticket Triage")
    print(f"  {datetime.today().strftime('%d %B %Y  %H:%M')}")
    print_divider("=")
    print()

    tickets = load_tickets(TICKETS_FILE)
    print(f"  Loaded {len(tickets)} tickets from {TICKETS_FILE}")
    print(f"  Model: {GROQ_MODEL}")
    print()

    results = []

    for i, ticket in enumerate(tickets, start=1):
        print(f"  Processing ticket {i}/{len(tickets)}: {ticket['ticket_id']}...", end=" ", flush=True)
        prompt = build_prompt(ticket)
        result = call_groq(prompt)
        results.append(result)
        print("done")

    print()

    for i, (ticket, result) in enumerate(zip(tickets, results), start=1):
        print_ticket_result(ticket, result, i, len(tickets))

    print_summary(tickets, results)

if __name__ == "__main__":
    main()

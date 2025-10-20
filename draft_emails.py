import pandas as pd

# Path to Excel
EXCEL_PATH = "master_contacts.xlsx"
OUTPUT_FILE = "email_drafts.txt"

# --- EMAIL TEMPLATE (no custom snippet) ---
def make_email(first_name, company, industry):
    subject = f"Quick call to discuss {industry} insights with UWO entrepreneur"
    body = f"""Subject: {subject}

Hi {first_name},

My name is Owen Clark. I’m a student at the University of Western Ontario and in the process of building DealScout AI, an AI-powered marketplace that connects buyers and sellers of small- and mid-sized companies. The goal is to make the early stages of M&A — growth, succession, and sale conversations — simpler and less time-consuming.

I’m reaching out to a few owners in {industry} around London to learn what challenges come up when people start thinking about growth, succession, or selling. Talking directly with business owners helps us validate assumptions and build something genuinely useful.

When I came across {company}, I wanted to reach out.

No need to be thinking about selling — I’d just really value your perspective as a business owner. Would you be open to a short 15–20 minute virtual chat with our team next week? I’d really value your perspective and can share more about the idea.

Best,
Owen Clark
CEO | DealScout AI
"""
    return body


# --- MAIN SCRIPT ---
def main():
    df = pd.read_excel(EXCEL_PATH)
    drafts = []

    for _, row in df.iterrows():
        first_name = row.get("First Name", "").strip()
        email = row.get("Email", "").strip()
        company = row.get("Company Name", "").strip()
        industry = row.get("Industry", "").strip()

        if not email:
            print(f"⚠️ Skipping missing email for {first_name}")
            continue

        email_text = make_email(first_name, company, industry)
        drafts.append(f"---\nTo: {email}\n\n{email_text}\n")

    # Save all drafts to a text file
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(drafts))

    print(f"✅ {len(drafts)} email drafts created in '{OUTPUT_FILE}'")


if __name__ == "__main__":
    main()

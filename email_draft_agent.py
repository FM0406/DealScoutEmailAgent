"""
FREE Email Draft Generator - Choose Your Method

METHOD 1: Generate Gmail Draft URLs (100% FREE, NO API NEEDED)
- Opens Gmail compose window with pre-filled content
- Most user-friendly option
- Works instantly, no setup

METHOD 2: Export to CSV for Gmail Mail Merge
- Use free Gmail add-on "Yet Another Mail Merge"
- Import CSV and send from Gmail

METHOD 3: Create .eml files
- Import into any email client (Gmail, Outlook, etc.)

Choose method below by setting METHOD variable
"""

import pandas as pd
import urllib.parse
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# =========================
# CONFIGURATION
# =========================
EXCEL_PATH = "master_contacts.xlsx"

# Choose method: "gmail_url", "csv_mailmerge", or "eml_files"
METHOD = "gmail_url"

# Sender info - each person changes this
SENDER_NAME = "Owen Clark"
SENDER_TITLE = "CEO | DealScout AI"

# =========================
# EMAIL TEMPLATE
# =========================
def create_email_content(first_name, company, industry):
    """Generate personalized email subject and body"""
    subject = f"Quick call to discuss {industry} insights with UWO entrepreneur"
    
    body = f"""Hi {first_name},

My name is {SENDER_NAME}. I'm a student at the University of Western Ontario and in the process of building DealScout AI, an AI-powered marketplace that connects buyers and sellers of small- and mid-sized companies. The goal is to make the early stages of M&A — growth, succession, and sale conversations — simpler and less time-consuming.

I'm reaching out to a few owners in {industry} around London to learn what challenges come up when people start thinking about growth, succession, or selling. Talking directly with business owners helps us validate assumptions and build something genuinely useful.

When I came across {company}, I wanted to reach out.

No need to be thinking about selling — I'd just really value your perspective as a business owner. Would you be open to a short 15–20 minute virtual chat with our team next week? I'd really value your perspective and can share more about the idea.

Best,
{SENDER_NAME}
{SENDER_TITLE}"""
    
    return subject, body

# =========================
# METHOD 1: GMAIL URL LINKS
# =========================
def generate_gmail_urls(df):
    """Generate clickable Gmail compose URLs - EASIEST & FREE"""
    print("\n📧 METHOD 1: Gmail Compose URLs")
    print("=" * 70)
    
    html_output = """<!DOCTYPE html>
<html>
<head>
    <title>Email Drafts - Click to Open</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { color: #1a73e8; }
        .email-card { 
            border: 1px solid #ddd; 
            padding: 20px; 
            margin: 15px 0; 
            border-radius: 5px; 
            background: #fafafa;
        }
        .email-card:hover { background: #f0f0f0; }
        .btn { 
            display: inline-block;
            background: #1a73e8; 
            color: white; 
            padding: 12px 24px; 
            text-decoration: none; 
            border-radius: 5px;
            font-weight: bold;
            margin-top: 10px;
        }
        .btn:hover { background: #1557b0; }
        .info { color: #666; font-size: 14px; }
        .preview { 
            background: white; 
            padding: 15px; 
            border-left: 3px solid #1a73e8; 
            margin: 10px 0;
            font-size: 13px;
            color: #333;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📧 Email Drafts Ready</h1>
        <p style="color: #666; font-size: 16px;">Click any button below to open Gmail with a pre-filled draft. Review and send!</p>
        <hr style="margin: 20px 0;">
"""
    
    for idx, row in df.iterrows():
        first_name = str(row.get("First Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        company = str(row.get("Company Name", "")).strip()
        industry = str(row.get("Industry", "")).strip()
        
        if not email or email == "nan":
            continue
        
        subject, body = create_email_content(first_name, company, industry)
        
        # Create Gmail URL
        gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&to={urllib.parse.quote(email)}&su={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"
        
        # Add to HTML
        html_output += f"""
        <div class="email-card">
            <h3>#{idx + 1} - {first_name} ({company})</h3>
            <div class="info">
                <strong>To:</strong> {email}<br>
                <strong>Industry:</strong> {industry}
            </div>
            <div class="preview">
                <strong>Subject:</strong> {subject}<br><br>
                {body[:200]}...
            </div>
            <a href="{gmail_url}" class="btn" target="_blank">✉️ Open Draft in Gmail</a>
        </div>
        """
    
    html_output += """
    </div>
</body>
</html>
"""
    
    # Save HTML file
    output_file = "email_drafts.html"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_output)
    
    print(f"✅ Created '{output_file}'")
    print(f"💡 Open this file in your browser and click any button to create drafts!")
    print(f"📊 Total drafts: {len([r for _, r in df.iterrows() if str(r.get('Email', '')).strip() and str(r.get('Email', '')).strip() != 'nan'])}")
    
    return output_file

# =========================
# METHOD 2: CSV FOR MAIL MERGE
# =========================
def generate_mail_merge_csv(df):
    """Generate CSV for Gmail Mail Merge add-on"""
    print("\n📧 METHOD 2: Mail Merge CSV")
    print("=" * 70)
    
    merge_data = []
    
    for idx, row in df.iterrows():
        first_name = str(row.get("First Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        company = str(row.get("Company Name", "")).strip()
        industry = str(row.get("Industry", "")).strip()
        
        if not email or email == "nan":
            continue
        
        subject, body = create_email_content(first_name, company, industry)
        
        merge_data.append({
            "Email": email,
            "First Name": first_name,
            "Company": company,
            "Industry": industry,
            "Subject": subject,
            "Message": body
        })
    
    # Save to CSV
    output_file = "mail_merge.csv"
    merge_df = pd.DataFrame(merge_data)
    merge_df.to_csv(output_file, index=False)
    
    print(f"✅ Created '{output_file}'")
    print(f"\n📝 NEXT STEPS:")
    print(f"1. Install 'Yet Another Mail Merge' (free Gmail add-on)")
    print(f"2. Open Gmail → Extensions → Yet Another Mail Merge")
    print(f"3. Import '{output_file}'")
    print(f"4. Review and send!")
    print(f"📊 Total contacts: {len(merge_data)}")
    
    return output_file

# =========================
# METHOD 3: .EML FILES
# =========================
def generate_eml_files(df):
    """Generate individual .eml files that can be imported into any email client"""
    print("\n📧 METHOD 3: .EML Files")
    print("=" * 70)
    
    # Create output folder
    output_folder = "email_drafts_eml"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    created_count = 0
    skipped_count = 0
    
    for idx, row in df.iterrows():
        first_name = str(row.get("First Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        company = str(row.get("Company Name", "")).strip()
        industry = str(row.get("Industry", "")).strip()
        
        if not email or email == "nan":
            print(f"⚠️  Row {idx + 2}: Skipping - no email for {first_name}")
            skipped_count += 1
            continue
        
        subject, body = create_email_content(first_name, company, industry)
        
        # Create email message
        msg = MIMEMultipart()
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        # Safe filename
        safe_name = f"{idx + 1:03d}_{first_name.replace(' ', '_')}_{company.replace(' ', '_')[:30]}.eml"
        safe_name = "".join(c for c in safe_name if c.isalnum() or c in ('_', '-', '.'))
        
        # Save .eml file
        filepath = os.path.join(output_folder, safe_name)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(msg.as_string())
        
        print(f"✅ Created: {safe_name}")
        created_count += 1
    
    print("\n" + "=" * 70)
    print(f"📊 SUMMARY")
    print(f"✅ Created: {created_count} .eml files")
    print(f"⚠️  Skipped: {skipped_count}")
    print(f"📁 Location: ./{output_folder}/")
    
    print(f"\n💡 HOW TO USE:")
    print(f"   GMAIL:")
    print(f"   1. Go to Gmail")
    print(f"   2. Drag and drop .eml files into compose window")
    print(f"   3. Or use 'Import mail & contacts' in Settings")
    
    print(f"\n   OUTLOOK:")
    print(f"   1. Double-click any .eml file")
    print(f"   2. Opens in Outlook as draft")
    print(f"   3. Review and send")
    
    print(f"\n   APPLE MAIL:")
    print(f"   1. Drag .eml files to Mail app")
    print(f"   2. Opens as draft")
    print(f"   3. Review and send")
    
    print("=" * 70)
    
    return output_folder

# =========================
# MAIN EXECUTION
# =========================
def main():
    print("\n" + "=" * 70)
    print("FREE EMAIL DRAFT GENERATOR")
    print("=" * 70)
    
    # Check if Excel exists
    if not os.path.exists(EXCEL_PATH):
        print(f"\n❌ ERROR: '{EXCEL_PATH}' not found!")
        print("Please make sure your Excel file is in the same folder.")
        return
    
    # Load Excel
    print(f"\n📂 Loading contacts from '{EXCEL_PATH}'...")
    try:
        df = pd.read_excel(EXCEL_PATH)
        print(f"✅ Loaded {len(df)} contacts")
    except Exception as e:
        print(f"❌ Failed to read Excel: {e}")
        return
    
    # Validate columns
    required_cols = ["First Name", "Email", "Company Name", "Industry"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        print(f"\n❌ Missing columns: {', '.join(missing)}")
        print(f"Required columns: {', '.join(required_cols)}")
        return
    
    # Execute selected method
    if METHOD == "gmail_url":
        generate_gmail_urls(df)
    elif METHOD == "csv_mailmerge":
        generate_mail_merge_csv(df)
    elif METHOD == "eml_files":
        generate_eml_files(df)
    else:
        print(f"❌ Invalid METHOD: '{METHOD}'")
        print("Valid options: 'gmail_url', 'csv_mailmerge', 'eml_files'")

if __name__ == "__main__":
    main()
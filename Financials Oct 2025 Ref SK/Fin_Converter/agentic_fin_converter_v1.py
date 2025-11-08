import os
import re
import pandas as pd
from docx import Document

# --- CONFIG ---
DATA_PATH = r"D:\Personal\Financials Oct 2025 Ref SK\Fin_Converter\Data\AccountStatements"
OUTPUT_SUMMARY = os.path.join(DATA_PATH, "account_summary.csv")
OUTPUT_TRANSACTIONS = os.path.join(DATA_PATH, "transactions_summary.csv")

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs])
        return text
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
        return ""

def parse_account_summary(text, year, month, file_name):
    checks_pattern = re.compile(r"Checks\s*(\d+)\s*checks\s*totaling\s*\$?([0-9,.\-()]+)", re.S)
    withdrawals_pattern = re.compile(r"Withdrawals\s*/\s*Debits\s*(\d+)\s*items\s*totaling\s*\$?([0-9,.\-()]+)", re.S)
    deposits_pattern = re.compile(r"Deposits\s*/\s*Credits\s*(\d+)\s*items\s*totaling\s*\$?([0-9,.\-()]+)", re.S)

    checks_match = checks_pattern.search(text)
    withdrawals_match = withdrawals_pattern.search(text)
    deposits_match = deposits_pattern.search(text)

    checks_count = checks_match.group(1) if checks_match else "N/A"
    checks_total = checks_match.group(2) if checks_match else "N/A"
    w_count = withdrawals_match.group(1) if withdrawals_match else "N/A"
    w_total = withdrawals_match.group(2) if withdrawals_match else "N/A"
    d_count = deposits_match.group(1) if deposits_match else "N/A"
    d_total = deposits_match.group(2) if deposits_match else "N/A"

    if not checks_match and not withdrawals_match and not deposits_match:
        return None

    return {
        "Year": year,
        "Month": month,
        "Checks Count": checks_count,
        "Checks Total": checks_total,
        "Withdrawals/Debits Total": w_total,
        "Deposits/Credits Total": d_total,
        "Criteria Met": "",
        "Source File": os.path.basename(file_name)
    }

def parse_transactions(text, year, month, type_label):
    transactions = []
    lines = text.split("\n")
    for line in lines:
        m = re.match(r"^(\d{2}/\d{2})\s+\$?([0-9,.]+)\s+(.+)", line.strip())
        if m:
            date, amount, desc = m.groups()
            # Heuristic to determine if the line is a debit or credit, based on the type_label
            if type_label == "Withdrawals / Debits" and "-" not in amount:
                transactions.append({
                    "Year": year,
                    "Month": month,
                    "Type": type_label,
                    "Date": date,
                    "Amount (USD)": amount,
                    "Description": desc.strip()
                })
            elif type_label == "Deposits / Credits" and "-" in amount:
                transactions.append({
                    "Year": year,
                    "Month": month,
                    "Type": type_label,
                    "Date": date,
                    "Amount (USD)": amount,
                    "Description": desc.strip()
                })

    return transactions

def extract_year_month(file_name):
    month_map = {
        "Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April",
        "May": "May", "Jun": "June", "Jul": "July", "Aug": "August",
        "Sept": "September", "Oct": "October", "Nov": "November", "Dec": "December"
    }
    m = re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sept|Oct|Nov|Dec)[a-z.]*\s*(\d{4})", file_name, re.I)
    if m:
        month, year = m.groups()
        return year, month_map.get(month[:3].title(), month)
    return "", ""

# --- MAIN ---
summary_records = []
transaction_records = []

for file in os.listdir(DATA_PATH):
    if file.lower().endswith(".docx"):
        file_path = os.path.join(DATA_PATH, file)
        print(f"📄 Processing {file} ...")
        text = extract_text_from_docx(file_path)
        year, month = extract_year_month(file)

        summary = parse_account_summary(text, year, month, file)
        if summary:
            summary_records.append(summary)

        for t_type in ["Withdrawals / Debits", "Deposits / Credits"]:
            transactions = parse_transactions(text, year, month, t_type)
            transaction_records.extend(transactions)

# --- SAVE CSVs ---
if summary_records:
    df_summary = pd.DataFrame(summary_records)
    df_summary.insert(0, "S. No.", range(1, len(df_summary) + 1))
    df_summary.to_csv(OUTPUT_SUMMARY, index=False)
    print(f"✅ Saved: {OUTPUT_SUMMARY} ({len(df_summary)} rows)")
else:
    print("⚠️ No account summaries extracted.")

if transaction_records:
    df_trans = pd.DataFrame(transaction_records)
    df_trans.to_csv(OUTPUT_TRANSACTIONS, index=False)
    print(f"✅ Saved: {OUTPUT_TRANSACTIONS} ({len(df_trans)} rows)")
else:
    print("⚠️ No transactions extracted.")

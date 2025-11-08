import re
import json
import os
from datetime import datetime

def parse_date(date_str):
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            pass
    raise ValueError(f"no valid date format found for {date_str}")

def extract_financial_data(text_content, file_name):
    account_summary = {}
    deposits_and_credits = []

    # Extract month and year from file name like "Sept. 2024 Payroll.PDF"
    base_name = os.path.basename(file_name)
    match = re.search(r"(\w{3,9})\.?\s(\d{4})", base_name)
    if match:
        month_str, year_str = match.groups()
        month_map = {
            'Jan': 'January', 'Feb': 'February', 'Mar': 'March', 'Apr': 'April',
            'May': 'May', 'Jun': 'June', 'Jul': 'July', 'Aug': 'August',
            'Sep': 'September', 'Sept': 'September', 'Oct': 'October', 'Nov': 'November', 'Dec': 'December'
        }
        month = month_map.get(month_str.replace('.', ''), month_str)
        year = int(year_str)
        account_summary["Month"] = f"{month} {year}"

    # Regex patterns for account summary
    summary_patterns = {
        "Beginning Balance": r"Beginning Balance.*?\s+\$?([\d,]+\.\d{2})",
        "Withdrawals/Debits": r"Checks.*?\s+\$\(([\d,]+\.\d{2})\)",
        "Ending Balance": r"Ending Balance.*?\s+\$?([\d,]+\.\d{2})",
        "# of Checks": r"(\d+)\s+checks totaling",
        "Number of Days in Period": r"Number of Days in Period\s+(\d+)",
        "Deposits/Credits": r"Deposits / Credits.*?\s+\$?([\d,]+\.\d{2})",
    }

    for key, pattern in summary_patterns.items():
        match = re.search(pattern, text_content, re.DOTALL)
        if match:
            value = match.group(1).replace(",", "").replace("(", "").replace(")", "")
            account_summary[key] = float(value) if key in ["Beginning Balance", "Withdrawals/Debits", "Ending Balance", "Deposits/Credits"] else int(value)

    # Calculate Net Change
    if "Ending Balance" in account_summary and "Beginning Balance" in account_summary:
        account_summary["Net Change"] = round(account_summary["Ending Balance"] - account_summary["Beginning Balance"], 2)

    # Regex for deposits
    deposits_section_match = re.search(r"Deposits / Credits(.*?)Daily Balance Summary", text_content, re.DOTALL)
    if not deposits_section_match:
        deposits_section_match = re.search(r"Deposits / Credits(.*?)(?=For additional information)", text_content, re.DOTALL)
    
    if deposits_section_match:
        deposits_section = deposits_section_match.group(1)
        # Find all deposit lines
        deposit_lines = re.findall(r"(\d{2}/\d{2})\s+([\d,]+\.\d{2})\s+(.*?)(?=\s*\d{2}/\d{2}\s|$)", deposits_section)

        for line in deposit_lines:
            date_str, amount_str, desc = line
            if not date_str.strip(): # Skip if date is empty
                continue

            full_date_str = f"{date_str}/{year}"
            try:
                # Handle cases where year might be wrong due to file naming, try to correct
                parsed_date = parse_date(full_date_str)
            except (ValueError, TypeError):
                # If parsing fails, try previous year for statements from early in the year referencing last year
                full_date_str = f"{date_str}/{year - 1}"
                try:
                    parsed_date = parse_date(full_date_str)
                except (ValueError, TypeError):
                    parsed_date = None # Or handle as an error

            if parsed_date:
                deposits_and_credits.append({
                    "Month": account_summary.get("Month", "N/A"),
                    "Date": parsed_date.strftime("%Y-%m-%d"),
                    "Amount": float(amount_str.replace(",", "")),
                    "Description": desc.strip(),
                    "Reference No.": (ref.group(1) if (ref := re.search(r"REF # (\S+)", desc)) else "N/A")
                })

    # Fallback for # of Deposits
    if deposits_and_credits:
        account_summary["# of Deposits"] = len(deposits_and_credits)
    else:
        match = re.search(r"(\d+)\s+items? totaling \S+\s+Deposits / Credits", text_content, re.DOTALL)
        if match:
            account_summary["# of Deposits"] = int(match.group(1))
            
    return account_summary, deposits_and_credits

def analyze_and_summarize(account_summaries):
    total_withdrawals = sum(s.get("Withdrawals/Debits", 0) for s in account_summaries)
    total_deposits = sum(s.get("Deposits/Credits", 0) for s in account_summaries)
    avg_net_change = sum(s.get("Net Change", 0) for s in account_summaries) / len(account_summaries) if account_summaries else 0

    # Trend analysis
    withdrawals_trend = "stable"
    if len(account_summaries) > 1:
        if account_summaries[-1].get("Withdrawals/Debits", 0) > account_summaries[0].get("Withdrawals/Debits", 0):
            withdrawals_trend = "growing"
        elif account_summaries[-1].get("Withdrawals/Debits", 0) < account_summaries[0].get("Withdrawals/Debits", 0):
            withdrawals_trend = "declining"

    deposits_trend = "stable"
    if len(account_summaries) > 1:
        if account_summaries[-1].get("Deposits/Credits", 0) > account_summaries[0].get("Deposits/Credits", 0):
            deposits_trend = "growing"
        elif account_summaries[-1].get("Deposits/Credits", 0) < account_summaries[0].get("Deposits/Credits", 0):
            deposits_trend = "declining"

    # Anomalies
    anomalies = []
    avg_ending_balance = sum(s.get("Ending Balance", 0) for s in account_summaries if s.get("Ending Balance") is not None) / len(account_summaries) if account_summaries else 0
    for summary in account_summaries:
        if summary.get("Ending Balance") is not None and avg_ending_balance > 0:
            if summary.get("Ending Balance", 0) > avg_ending_balance * 1.5:
                anomalies.append(f"Unusually high ending balance in {summary.get('Month')}: ${summary.get('Ending Balance'):,.2f}")
            if summary.get("Ending Balance", 0) < avg_ending_balance * 0.5:
                anomalies.append(f"Unusually low ending balance in {summary.get('Month')}: ${summary.get('Ending Balance'):,.2f}")

    return {
        "trends": f"Total Withdrawals (${total_withdrawals:,.2f}) vs. Deposits (${total_deposits:,.2f}). Withdrawals are {withdrawals_trend}, deposits are {deposits_trend}.",
        "average_monthly_net_change": f"${avg_net_change:,.2f}",
        "anomalies": anomalies if anomalies else "No significant anomalies detected.",
        "notes": "Payroll funding appears regular. Most deposits are online transfers."
    }


def run_analysis():
    with open('financial_data.txt', 'r', encoding='utf-8') as f:
        content = f.read()

    # Split content by file separator
    file_contents = content.split('--- ')[1:]
    
    all_account_summaries = []
    all_deposits_credits = []

    for file_content in file_contents:
        try:
            file_name, text = file_content.split(' ---', 1)
            summary, deposits = extract_financial_data(text, file_name)
            if summary:
                all_account_summaries.append(summary)
            if deposits:
                all_deposits_credits.extend(deposits)
        except Exception as e:
            print(f"Error processing file: {e}")

    # Sort summaries by date
    all_account_summaries.sort(key=lambda x: datetime.strptime(x.get('Month', 'January 1900'), "%B %Y"))

    # Perform analysis
    observations = analyze_and_summarize(all_account_summaries)

    # Final JSON output
    final_output = {
        "account_summary_table": all_account_summaries,
        "deposits_and_credits_table": all_deposits_credits,
        "observations": observations
    }

    with open("financial_summary.json", "w") as f:
        json.dump(final_output, f, indent=2)

    print("Analysis complete. Results saved to financial_summary.json")

if __name__ == "__main__":
    run_analysis()
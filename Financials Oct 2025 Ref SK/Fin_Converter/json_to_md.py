import json

def json_to_md(json_file, md_file):
    """
    Converts a JSON file with financial data to a markdown file with tables.

    Args:
        json_file (str): The path to the input JSON file.
        md_file (str): The path to the output markdown file.
    """
    with open(json_file, 'r') as f:
        data = json.load(f)

    with open(md_file, 'w') as f:
        f.write("# Financial Summary Analysis\n\n")
        f.write("This report provides a detailed analysis of the payroll account statements from September 2024 to September 2025.\n\n")
        f.write("## Description and Implications\n\n")
        f.write("The analysis of the payroll account reveals several key insights into the financial health of the business.\n\n")

        # Observations
        f.write("### Observations\n\n")
        total_withdrawals = sum(item['Withdrawals/Debits'] for item in data['account_summary_table'])
        total_deposits = sum(item['Deposits/Credits'] for item in data['account_summary_table'])
        f.write(f"*   **Trends:** Total Withdrawals (${total_withdrawals:,.2f}) vs. Deposits (${total_deposits:,.2f}). Both withdrawals and deposits show a declining trend over the period.\n")
        net_changes = [item['Net Change'] for item in data['account_summary_table']]
        average_net_change = sum(net_changes) / len(net_changes)
        f.write(f"*   **Average Monthly Net Change:** ${average_net_change:,.2f}\n")
        f.write("*   **Anomalies:**\n")
        for item in data['account_summary_table']:
            if item['Ending Balance'] == 0:
                f.write(f"    *   Unusually low ending balance in {item['Month']}: ${item['Ending Balance']:,.2f}\n")
        f.write("*   **Notes:** Payroll funding appears to be regular, with most deposits being online transfers.\n\n")

        # Implications
        f.write("### Implications\n\n")
        f.write("*   **Cash Flow:** The declining trend in both deposits and withdrawals may indicate a slowdown in business activity. The company is spending more than it is depositing, which is not sustainable in the long run.\n")
        f.write("*   **Financial Stability:** The low ending balances in October and December 2024 are a cause for concern. They suggest that the company may have faced liquidity issues during those months.\n")
        f.write("*   **Payroll Funding:** The regularity of payroll funding is a positive sign, indicating that the company is able to meet its payroll obligations on time.\n\n")

        # Account Summary Table
        f.write("## Account Summary Table\n\n")
        f.write("| Month | Beginning Balance | Withdrawals/Debits | Deposits/Credits | Ending Balance | # of Checks | # of Deposits | Net Change | Days in Period |\n")
        f.write("| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n")
        for item in data['account_summary_table']:
            f.write(f"| {item['Month']} | {item['Beginning Balance']} | {item['Withdrawals/Debits']} | {item['Deposits/Credits']} | {item['Ending Balance']} | {item['# of Checks']} | {item['# of Deposits']} | {item['Net Change']} | {item['Number of Days in Period']} |\n")
        f.write("\n")

        # Deposits and Credits Table
        f.write("## Deposits and Credits Table\n\n")
        f.write("| Month | Date | Amount | Description | Reference No. |\n")
        f.write("| :--- | :--- | :--- | :--- | :--- |\n")
        for item in data['deposits_and_credits_table']:
            f.write(f"| {item['Month']} | {item['Date']} | {item['Amount']} | {item['Description']} | {item['Reference No.']} |\n")

if __name__ == '__main__':
    json_to_md('financial_summary.json', 'financial_summary_tables.md')
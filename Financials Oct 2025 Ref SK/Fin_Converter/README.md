# Financial Report Generator

## Description

This project is a financial report generator that analyzes payroll data and creates a summary report in JSON, Markdown, and PDF formats.

## Installation

1.  Clone the repository:
    ```bash
    git clone https://github.com/your-username/financial-report-generator.git
    ```
2.  Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Place your payroll data files in the `Data/PayrollStatements` directory.
2.  Run the `financial_analyzer.py` script to generate the financial summary in JSON format:
    ```bash
    python financial_analyzer.py
    ```
3.  Run the `json_to_md.py` script to convert the JSON summary to a Markdown table:
    ```bash
    python json_to_md.py
    ```
4.  The generated reports will be available in the root directory as `financial_summary.json`, `financial_summary_tables.md`, and `financial_summary_tables.pdf`.

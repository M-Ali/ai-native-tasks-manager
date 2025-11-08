"""
document_reader_mcp.py
-----------------------------------
MCP server for reading and extracting text from financial documents:
- PDF (via pdfplumber, camelot, OCR)
- DOCX
- TXT
"""

import os
import re
import pandas as pd
import pdfplumber
import camelot
from pdf2image import convert_from_path
import pytesseract
from docx import Document

# ==============================
# Core Class: DocumentReaderMCP
# ==============================
class DocumentReaderMCP:
    def __init__(self, base_folder):
        self.base_folder = base_folder
        self.date_pattern = re.compile(r'^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}')

    def _parse_lines(self, lines, file_name, method):
        data = []
        current_type = None
        for line in lines:
            if any(x in line.lower() for x in ["withdrawals", "debits"]):
                current_type = "Debit"
                continue
            elif any(x in line.lower() for x in ["deposits", "credits"]):
                current_type = "Credit"
                continue
            elif "summary" in line.lower():
                current_type = None
                continue

            if not current_type or not self.date_pattern.match(line):
                continue

            parts = re.split(r'\s{2,}', line)
            date = parts[0]
            description = parts[1] if len(parts) > 1 else ""
            amount = parts[-1] if len(parts) > 2 else ""

            try:
                d = pd.to_datetime(date, format="%d %b %Y")
                year, month = d.year, d.month_name()
            except:
                year, month = "", ""

            data.append({
                "Date": date,
                "Description": description,
                "Amount": amount,
                "Type": current_type,
                "Year": year,
                "Month": month,
                "Source File": file_name,
                "Method": method
            })
        return data

    def _clean_amount(self, df):
        df["Amount"] = (
            df["Amount"].astype(str)
            .str.replace(",", "")
            .str.replace("CR", "")
            .str.replace("DR", "")
            .str.strip()
        )
        return df

    def read_documents(self):
        pdf_results, camelot_results, ocr_results, docx_results = [], [], [], []

        for root, _, files in os.walk(self.base_folder):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                ext = os.path.splitext(file_name)[-1].lower()
                print(f"\n📘 Reading: {file_name}")

                # 1️⃣ PDF via pdfplumber
                if ext == ".pdf":
                    try:
                        with pdfplumber.open(file_path) as pdf:
                            text_lines = []
                            for page in pdf.pages:
                                text_lines += [l.strip() for l in page.extract_text().splitlines() if l.strip()]
                            pdf_results.extend(self._parse_lines(text_lines, file_name, "pdfplumber"))
                            print(f"✅ pdfplumber extracted {len(text_lines)} lines")
                    except Exception as e:
                        print(f"⚠️ pdfplumber failed: {e}")

                    # 2️⃣ Camelot
                    try:
                        tables = camelot.read_pdf(file_path, pages="all", flavor="stream")
                        for t in tables:
                            df = t.df
                            df["Source File"] = file_name
                            df["Method"] = "camelot"
                            camelot_results.append(df)
                        print(f"✅ camelot extracted {len(tables)} tables")
                    except Exception as e:
                        print(f"⚠️ Camelot failed: {e}")

                    # 3️⃣ OCR fallback
                    try:
                        images = convert_from_path(file_path, dpi=200)
                        ocr_text = []
                        for img in images:
                            ocr_text += pytesseract.image_to_string(img).splitlines()
                        ocr_results.extend(self._parse_lines(ocr_text, file_name, "ocr"))
                        print(f"✅ OCR extracted {len(ocr_text)} lines")
                    except Exception as e:
                        print(f"⚠️ OCR failed: {e}")

                # 4️⃣ DOCX
                elif ext == ".docx":
                    try:
                        doc = Document(file_path)
                        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                        docx_results.extend(self._parse_lines(lines, file_name, "docx"))
                        print(f"✅ DOCX extracted {len(lines)} lines")
                    except Exception as e:
                        print(f"⚠️ DOCX failed: {e}")

                # 5️⃣ TXT
                elif ext == ".txt":
                    with open(file_path, "r", encoding="utf-8") as f:
                        lines = [l.strip() for l in f if l.strip()]
                    docx_results.extend(self._parse_lines(lines, file_name, "txt"))
                    print(f"✅ TXT extracted {len(lines)} lines")

        # Merge DataFrames
        dfs = []
        if pdf_results:
            dfs.append(pd.DataFrame(pdf_results))
        if ocr_results:
            dfs.append(pd.DataFrame(ocr_results))
        if docx_results:
            dfs.append(pd.DataFrame(docx_results))
        if camelot_results:
            dfs.append(pd.concat(camelot_results, ignore_index=True))

        if not dfs:
            print("⚠️ No data extracted.")
            return pd.DataFrame()

        final_df = pd.concat(dfs, ignore_index=True)
        final_df = self._clean_amount(final_df)
        print(f"\n✅ Total combined rows: {len(final_df)}")
        return final_df

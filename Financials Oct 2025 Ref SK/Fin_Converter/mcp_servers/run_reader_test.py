from mcp_servers.document_reader_mcp import DocumentReaderMCP

folder = r"D:\Personal\Financials Oct 2025 Ref SK\AgenticFin\Data\AccountStatements"
reader = DocumentReaderMCP(folder)
df = reader.read_documents()
df.to_csv("D:/Personal/Financials Oct 2025 Ref SK/AgenticFin/Outputs/account_reader_results.csv", index=False, encoding="utf-8-sig")
print(df.head(10))

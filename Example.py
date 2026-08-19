from Codal_statement.Codal import *


# ------------------------------------------------------------
# 1. Get raw financial data from Codal
# ------------------------------------------------------------
# Retrieves the raw financial statement data from Codal
# based on the selected company, statement, and filters.

all_datasource = get_financial_report(
    symbol="زاگرس",
    statement_name="صورت سود و زیان",
    length=12,
    from_date="1400/01/01",
    to_date="1405/01/01",
    mains=True,
    consolidatable=True,
    NotConsolidatable=True,
    NotAudited=False,
    Audited=True,
)


# ------------------------------------------------------------
# 2. Convert raw data to DataFrame
# ------------------------------------------------------------
# Converts the raw Codal data into a Pandas DataFrame.

df = report_to_dataframe(all_datasource)


# ------------------------------------------------------------
# 3. Export financial statements to Excel
# ------------------------------------------------------------
# Converts the financial statements into an Excel file.

excel = report_to_excel(
    all_datasource,
    path="E:/Report.xlsx"
)


# ------------------------------------------------------------
# 4. Download reports and attachments
# ------------------------------------------------------------
# Downloads the available PDF, Excel, and attachment files
# related to the financial reports.

download_report(
    all_datasource,
    "E:/codal_files"
)
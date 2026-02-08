import pandas as pd
import requests

month = input("Enter month (MM): ")
year = input("Enter year (YY): ")

base = "https://www.axismf.com/cms/sites/default/files/Statutory/Monthly%20Portfolio-"

found = False

for day in ["31", "30", "29", "28"]:
    date_part = f"{day}%20{month}%20{year}"
    url = base + date_part + ".xlsx"

    print("Trying:", url)

    r = requests.get(url, timeout=15)
    if r.status_code == 200:
        filename = f"Monthly_Portfolio_{day}_{month}_{year}.xlsx"
        with open(filename, "wb") as f:
            f.write(r.content)

        print("Downloaded:", filename)
        found = True
        break


if not found:
    print("No file found for that month/year.")



filename = "Monthly_Portfolio_31_12_25.xlsx"  

df = pd.read_excel(filename)
#print(df.head())

xl = pd.ExcelFile(filename)


from openpyxl import load_workbook

wb = load_workbook(filename, data_only=True)

consolidated_data = []

for sheet_name in wb.sheetnames:
    if sheet_name.lower() not in ["index", "consolidated"]:

        ws = wb[sheet_name]
        scheme_name = sheet_name

        report_date = None
        for row in ws.iter_rows(min_row=1, max_row=5): # date is in B3, so we check first 5 rows for more robustness
            for cell in row:
                if cell.value and "as on" in str(cell.value).lower():
                    report_date = str(cell.value)
                    break

        for row in ws.iter_rows():

            name_cell = row[1]   # column B

            if not name_cell.value:
                continue
            if name_cell.font.bold:
                continue

            instrument = str(name_cell.value).strip()

            isin = row[2].value      # column C
            market_value = row[5].value   # column F
            portfolio_pct = row[6].value  # column G

            consolidated_data.append({
                "Scheme Name": scheme_name,
                "Instrument Name": instrument,
                "ISIN": isin,
                "Market Value": market_value,
                "pct Portfolio": portfolio_pct,
                "Reporting Date": report_date
            })


#instrument classification based on name

def classify_instrument(name):
    x = str(name).lower()

    equity_keywords = [
        "ltd", "limited", "equity", "shares", "plc", "corp"
    ]

    debt_keywords = [
        "bond", "debenture", "ncd", "treasury", "g-sec",
        "government", "commercial paper", "cp",
        "certificate of deposit", "cd", "note", "repo"
    ]

    if any(k in x for k in equity_keywords):
        return "Equity"

    elif any(k in x for k in debt_keywords):
        return "Debt"

    else:
        return "Other"

consolidated_df = pd.DataFrame(consolidated_data)
consolidated_df ["AMC name"]= "Axis Bank"
consolidated_df.insert(0, "AMC name", consolidated_df.pop("AMC name"))  # move AMC name to first column
consolidated_df["Instrument Type"] = consolidated_df["Instrument Name"].apply(classify_instrument)
consolidated_df.to_excel("consolidated_data.xlsx", index=False) 

print("success")
        
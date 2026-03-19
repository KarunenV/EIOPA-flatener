from pathlib import Path
import pandas as pd

# 1) path setup
input_dir = Path("Input")
output_file = Path("merged_all.xlsx")

# simple currency mapping from file name
currency_map = {
    "euro": "EUR",
    "united kingdom": "GBP",
    "united states": "USD"
}

rows = []

for xlsx_path in sorted(input_dir.glob("*.xlsx")):
    print("Loading", xlsx_path.name)
    excel = pd.ExcelFile(xlsx_path, engine="openpyxl")

    name_key = xlsx_path.stem.lower()
    currency = next((v for k, v in currency_map.items() if k in name_key), None)

    print("  detected currency:", currency if currency else "None")

    # if currency is not detected, we stop processing this file, as we don't want to mix data with unknown currency
    if currency is None:
        print("  Warning: Currency not detected from file name")
        exit(0)

    for sheet_name in excel.sheet_names:
        print("  sheet", sheet_name)
        df = pd.read_excel(excel, sheet_name=sheet_name, engine="openpyxl")

        if "with" in sheet_name.lower():
            curve_type = "RFR_spot_with_VA"
        else:
            curve_type = "RFR_spot_no_VA"

        if "manual" in sheet_name.lower():
            manual_or_rss = "Manual"
        else:            
            manual_or_rss = "RSS"

        # For each column (skipping first column), each value is one output row
        for col in df.columns[1:]:
            date_value = col

            tenor = 1
            for cell in df[col][8:]:  # skip first 9 rows which are not data
                if pd.isna(cell):
                    tenor += 1
                    continue

                rows.append({
                    "Curve": curve_type,
                    "Currency": currency,
                    "Date": date_value,
                    "Tenor": tenor,
                    "Yield": cell,
                    "ManualorRSS": manual_or_rss
                })
                tenor += 1

if not rows:
    raise RuntimeError("No data rows found in Input/*.xlsx")

merged = pd.DataFrame(rows)


# write output
merged.to_excel(output_file, index=False, engine="openpyxl")
print("Merged file written to", output_file)
import pandas as pd

def add_pivot_with_mapping(writer, df, mapping_df):
    pd.set_option('future.no_silent_downcasting', True)
    workbook = writer.book
    pivot_sheet_name = "Sheet2"
    if pivot_sheet_name in writer.sheets:
        pivot_ws = writer.sheets[pivot_sheet_name]
    else:
        pivot_ws = workbook.add_worksheet(pivot_sheet_name)
        writer.sheets[pivot_sheet_name] = pivot_ws

    # Formatting
    header_fmt = workbook.add_format({'bold': True})
    number_fmt = workbook.add_format({'num_format':'#,##0.00'})
    yellow_fmt = workbook.add_format({'bg_color':'#FFFF00'})

    numeric_cols = [
        "Net Debit (NC)",
        "Commission (NC)",
        "Markup (NC)",
        "Scheme Fees (NC)",
        "Interchange (NC)",
        "Gross Credit (GC)",
        "Gross Debit (GC)"
    ]

    for col in numeric_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)  # remove thousands separators
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Compute pivot in Python
    df["Total cost calc"] = df.apply(
        lambda r: -r["Net Debit (NC)"] if r["Type"]=="Fee"
        else -r[["Commission (NC)","Markup (NC)","Scheme Fees (NC)","Interchange (NC)"]].fillna(0).sum(),
        axis=1
    )

    df["Total rev calc"] = df["Gross Credit (GC)"].fillna(0) - df["Gross Debit (GC)"].fillna(0)

    storeid_map = dict(zip(mapping_df["Store ID"], mapping_df["Description"]))
    df["StoreID calc"] = df["Store"].map(storeid_map).fillna("#N/B")

    pivot_df = df.groupby("StoreID calc")[["Total cost calc","Total rev calc"]].sum().reset_index()
    pivot_df.rename(columns={"StoreID calc":"StoreID","Total cost calc":"Total cost","Total rev calc":"Total rev"}, inplace=True)

    pivot_df = pd.concat([
        pivot_df[pivot_df["StoreID"] != "#N/B"],
        pivot_df[pivot_df["StoreID"] == "#N/B"]
    ], ignore_index=True)
        
    # Write pivot table headers
    for col_idx, col_name in enumerate(pivot_df.columns):
        pivot_ws.write(2, col_idx, col_name, header_fmt)

    # Write pivot values
    for row_idx, row in pivot_df.iterrows():
        for col_idx, value in enumerate(row):
            pivot_ws.write(row_idx + 3, col_idx, value, number_fmt)

    n_rows = len(pivot_df)
    start_col = len(pivot_df.columns) + 1

    # Mapping headers next to pivot
    mapping_headers = ["CC","CC EUR","GL","GL EUR"]
    for i, col_name in enumerate(mapping_headers):
        pivot_ws.write(2, start_col + i, col_name, yellow_fmt)

    # Write formulas for mapping columns
    for i in range(n_rows):
        excel_row = i + 4
        # CC
        pivot_ws.write_formula(
            i + 3, start_col,
            f'=IFERROR(INDEX(\'Adyen Stores Mapping\'!$I$2:$I${len(mapping_df)+1},MATCH(A{excel_row},\'Adyen Stores Mapping\'!$A$2:$A${len(mapping_df)+1},0)),"ALG")',
            yellow_fmt
        )
        # CC EUR
        pivot_ws.write_formula(
            i + 3, start_col+1, f"=B{excel_row}", yellow_fmt
        )
        # GL
        pivot_ws.write_formula(
            i + 3, start_col+2,
            f'=IFERROR(INDEX(\'Adyen Stores Mapping\'!$H$2:$H${len(mapping_df)+1},MATCH(A{excel_row},\'Adyen Stores Mapping\'!$A$2:$A${len(mapping_df)+1},0)),"")',
            yellow_fmt
        )
        # GL EUR
        pivot_ws.write_formula(
            i + 3, start_col+3, f"=C{excel_row}", yellow_fmt
        )
    
    first_data_row_excel = 4
    last_data_row_excel = 3 + n_rows  # because data starts on row 4

    totals_row_excel = last_data_row_excel + 2  # one empty row
    totals_row_index = totals_row_excel - 1     # zero-based index

    # Sum Total Cost
    pivot_ws.write_formula(
        totals_row_index,
        1,
        f"=SUM(B{first_data_row_excel}:B{last_data_row_excel})",
        header_fmt
    )

    # Sum Total Rev
    pivot_ws.write_formula(
        totals_row_index,
        2,
        f"=SUM(C{first_data_row_excel}:C{last_data_row_excel})",
        header_fmt
    )

    pivot_ws.write_formula(
        totals_row_index + 1,
        2,
        f"=SUM(B{totals_row_excel}:C{totals_row_excel})",
        number_fmt
    )

    # -----------------------------
    # CC EUR Combined SUM
    # -----------------------------

    cc_sum_row_excel = last_data_row_excel + 2
    cc_sum_row_index = cc_sum_row_excel

    pivot_ws.write_formula(
        cc_sum_row_index,
        5,
        f"=SUM(F{first_data_row_excel}:F{last_data_row_excel},H{first_data_row_excel}:H{last_data_row_excel})"
    )


import pandas as pd
import io

def create_excel_with_formulas(filename, df, mapping_df):

    output = io.BytesIO()

    # Add formula columns
    df["StoreID"] = ""
    df["Total cost"] = ""
    df["Total rev"] = ""

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

        df.to_excel(writer, sheet_name=filename, index=False)
        mapping_df.to_excel(writer, sheet_name="Adyen Stores Mapping", index=False)

        worksheet = writer.sheets[filename]

        storeid_col = df.columns.get_loc("StoreID")
        totalcost_col = df.columns.get_loc("Total cost")
        totalrev_col = df.columns.get_loc("Total rev")

        for row in range(len(df)):
            excel_row = row + 2

            worksheet.write_formula(
                row+1, storeid_col,
                f"=INDEX('Adyen Stores Mapping'!$A$2:$A${len(mapping_df)+1},"
                f"MATCH(X{excel_row},'Adyen Stores Mapping'!$G$2:$G${len(mapping_df)+1},0))"
            )

            if df.iloc[row]["Type"] == "Fee":
                worksheet.write_formula(
                    row+1, totalcost_col,
                    f"=-SUM(O{excel_row})"
                )
            else:
                worksheet.write_formula(
                    row+1, totalcost_col,
                    f"=-SUM(Q{excel_row}:T{excel_row})"
                )

            worksheet.write_formula(
                row+1, totalrev_col,
                f"=L{excel_row}-K{excel_row}"
            )

        # Add pivot
        add_pivot_with_mapping(writer, df, mapping_df)

    output.seek(0)

    return output

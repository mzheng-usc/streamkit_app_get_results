import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta
from openpyxl.styles import numbers

from merge_excel import merge_excel_data  # we'll extract your logic into this function

st.set_page_config(page_title="📆 Full-Day Excel Merger", layout="centered")
st.title("🧮 Merge Excel Files by Date & Group ID")

# Upload inputs
table1_file = st.file_uploader("📄 Upload Table 1 (e.g. 3PM to midnight)", type="xlsx", key="file1")
combined_table_file = st.file_uploader("📄 Upload Combined Table (both periods)", type="xlsx", key="file2")

# Date input
col1, col2 = st.columns(2)
with col1:
    first_date = st.date_input("📆 First date (e.g., 2025-05-24)", value=datetime(2025, 5, 24))
with col2:
    second_date = st.date_input("📆 Second date (e.g., 2025-05-25)", value=datetime(2025, 5, 25))

run_merge = st.button("🔁 Run Merge")

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, na_rep='N/A')
        worksheet = writer.sheets['Sheet1']

        # Apply text format to avoid scientific notation
        id_columns = ['渠道ID', 'Ad Group ID', 'Ad ID']
        for col_idx, col_name in enumerate(df.columns, 1):
            if col_name in id_columns:
                for row_idx in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.number_format = '@'
                    if isinstance(cell.value, (int, float)):
                        cell.value = str(int(cell.value))
    return output.getvalue()

if run_merge and table1_file and combined_table_file:
    try:
        table1_path = table1_file
        combined_table_path = combined_table_file

        # Read table1 to infer target columns
        table1_df = pd.read_excel(table1_path)
        target_col_names = ["投放花费", '应用设备激活数', '付费用户数(首日)', 'd0']
        target_col_indices = [list(table1_df.columns).index(x) for x in target_col_names if x in table1_df.columns]
        target_columns = sorted(list(set(list(range(12)) + target_col_indices)))

        result_df = merge_excel_data(
            table1_path,
            combined_table_path,
            target_columns=target_columns,
            first_date=str(first_date),
            second_date=str(second_date),
            perform_sanity_check=False
        )

        st.success("✅ Merge completed successfully!")
        st.write("Preview of merged data:")
        st.dataframe(result_df.head(10))

        # Format output filename based on second_date
        output_filename = second_date.strftime("%m%d") + "_results.xlsx"
        
        st.download_button(
            "📥 Download Merged Excel File",
            data=to_excel_bytes(result_df),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Merge failed: {e}")

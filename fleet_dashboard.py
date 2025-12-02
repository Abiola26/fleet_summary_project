import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

st.set_page_config(page_title="Fleet Summary Tool", layout="centered")
st.title("🚚 Fleet Summary Tool")

st.sidebar.header("Select Functionality")
option = st.sidebar.radio("Choose an option:", ["📁 CSV Summary & Subtotals", "📂 Combine Multiple Excel Files"])

# ========== FUNCTION TO FORMAT & EXPORT EXCEL ========== #
def style_and_export_to_excel(df):
    output = BytesIO()
    df.to_excel(output, index=False, sheet_name='Summary')
    
    # Load workbook for styling
    wb = load_workbook(output)
    ws = wb.active

    # Header styling
    header_fill = PatternFill(start_color="007acc", end_color="007acc", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    # Bold subtotal rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if row[1].value == "Subtotal":
            for cell in row:
                cell.font = Font(bold=True)

    # Format amount column
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        cell = row[3]  # 'Total Amount (₦)' is column index 3 (zero-based)
        if isinstance(cell.value, (int, float)):
            cell.number_format = '"₦"#,##0.00'

    # Save styled Excel to BytesIO
    final_output = BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

# ========== FUNCTION 1: CSV File Upload and Summary ========== #
if option == "📁 CSV Summary & Subtotals":
    st.subheader("📥 Upload CSV File")
    csv_file = st.file_uploader("Upload a CSV file", type=["csv"], key="csv")

    if csv_file:
        try:
            df = pd.read_csv(csv_file)
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df['OnlyDate'] = df['Date'].dt.date
            df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)

            grouped = df.groupby(['OnlyDate', 'Fleet']).agg(
                TotalAmount=('Amount', 'sum'),
                FleetCount=('Fleet', 'count')
            ).reset_index()

            # Build formatted DataFrame
            formatted_rows = []
            for date, group in grouped.groupby('OnlyDate'):
                subtotal_amount = group['TotalAmount'].sum()
                subtotal_count = group['FleetCount'].sum()

                for _, row in group.iterrows():
                    formatted_rows.append({
                        "Date": row['OnlyDate'],
                        "Fleet": row['Fleet'],
                        "Fleet Count": row['FleetCount'],
                        "Total Amount (₦)": f"{row['TotalAmount']:,.2f}"
                    })

                formatted_rows.append({
                    "Date": date,
                    "Fleet": "Subtotal",
                    "Fleet Count": subtotal_count,
                    "Total Amount (₦)": f"{subtotal_amount:,.2f}"
                })

            result_df = pd.DataFrame(formatted_rows)
            st.success("✅ CSV processed successfully!")
            st.dataframe(result_df)

            # Download button
            styled_excel = style_and_export_to_excel(result_df)
            st.download_button(
                label="📥 Download Summary as Excel",
                data=styled_excel,
                file_name="Fleet_Summary_Subtotal.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            st.error(f"❌ Error processing CSV: {e}")

# ========== FUNCTION 2: Combine Multiple Excel Files ========== #
elif option == "📂 Combine Multiple Excel Files":
    st.subheader("📥 Upload Excel Files")
    uploaded_files = st.file_uploader("Upload multiple Excel files", type=["xlsx"], accept_multiple_files=True, key="excel")

    if uploaded_files:
        all_dataframes = []

        for file in uploaded_files:
            try:
                df = pd.read_excel(file, sheet_name='Sheet1')
                df['Fleet Count'] = pd.to_numeric(df['Fleet Count'], errors='coerce')
                df['Total Amount (₦)'] = pd.to_numeric(df['Total Amount (₦)'], errors='coerce')
                df['Fleet'] = df['Fleet'].astype(str)
                all_dataframes.append(df)
            except Exception as e:
                st.error(f"❌ Error in file `{file.name}`: {e}")

        if all_dataframes:
            combined_df = pd.concat(all_dataframes, ignore_index=True)

            summary = combined_df.groupby('Fleet').agg({
                'Fleet Count': 'sum',
                'Total Amount (₦)': 'sum'
            }).reset_index().sort_values(by='Fleet')

            st.success("✅ Files combined and summarized successfully!")
            st.subheader("🚛 Combined Fleet Summary")
            st.dataframe(summary)

            @st.cache_data
            def convert_df_to_excel(dataframe):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    dataframe.to_excel(writer, index=False, sheet_name='FleetSummary')
                return output.getvalue()

            excel_data = convert_df_to_excel(summary)
            st.download_button(
                label="📥 Download Combined Summary as Excel",
                data=excel_data,
                file_name='Combined_Fleet_Summary.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

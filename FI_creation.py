import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Forecast Item Id Generator", layout="centered")
st.title("📊 Forecast Item Id Generator")

# File upload section
st.header("Upload Files")
data_manager_file = st.file_uploader("Upload Data Manager File (Excel)", type=["xlsx"])
user_request_file = st.file_uploader("Upload User Request File (Excel)", type=["xlsx"])

if data_manager_file and user_request_file:
    try:
        # Read files
        data_manager = pd.read_excel(data_manager_file, dtype=str).fillna('')
        user_request = pd.read_excel(user_request_file, dtype=str).fillna('')

        tactical_rows = []
        operational_rows = []

        today_str = datetime.today().strftime("%m/%d/%Y")
        
        # ✅ Corrected to include seconds
        end_eff_datetime = datetime(2998, 12, 31, 23, 59, 59)

        for _, req in user_request.iterrows():
            filtered = data_manager.copy()

            if req.get("ProductId", "").strip():
                filtered = filtered[filtered["ProductId"] == req["ProductId"]]

            if req.get("SCHSALESCHANNELCD", "").strip():
                filtered = filtered[filtered["SCHSALESCHANNELCD"] == req["SCHSALESCHANNELCD"]]

            if req.get("LOCATIONID", "").strip():
                filtered = filtered[filtered["LOCATIONID"] == req["LOCATIONID"]]

            if req.get("IntegratedPLanningAccountId", "").strip():
                val = req["IntegratedPLanningAccountId"].strip()
                filtered = filtered[filtered["IntegratedPLanningAccountId"].str.startswith(val, na=False)]

            filtered = filtered.copy()
            filtered["ForecastItemId"] = filtered["ForecastItemId"].apply(
                lambda x: x.replace(req["ProductId"], req["New SKU"])
            )
            filtered["ProductId"] = req["New SKU"]
            filtered["SOURCE"] = f"Manual Tarun Kumar {today_str}"
            filtered["ENDEFF"] = end_eff_datetime  # ✅ Correct datetime

            output_cols = [
                "ForecastItemId",
                "ProductId",
                "CUSTOMERID",
                "LOCATIONID",
                "ForecastItemType",
                "SOURCE",
                "ENDEFF"
            ]

            tactical_rows.append(filtered[filtered["ForecastItemType"] == "Tactical"][output_cols])
            operational_rows.append(filtered[filtered["ForecastItemType"] == "Operational"][output_cols])

        tactical_output = pd.concat(tactical_rows, ignore_index=True)
        operational_output = pd.concat(operational_rows, ignore_index=True)

        st.success("✅ Processing Complete")

        # Save tactical file
        tactical_bytes = io.BytesIO()
        with pd.ExcelWriter(tactical_bytes, engine='xlsxwriter') as writer:
            tactical_output.to_excel(writer, index=False, sheet_name='Tactical')
            workbook = writer.book
            worksheet = writer.sheets['Tactical']
            # ✅ Format for Excel display: 12/31/2998 23:59
            date_format = workbook.add_format({'num_format': 'mm/dd/yyyy hh:mm'})

            for i, col in enumerate(tactical_output.columns):
                max_len = max(tactical_output[col].astype(str).map(len).max(), len(col)) + 2
                if col == "ENDEFF":
                    worksheet.set_column(i, i, max_len, date_format)
                else:
                    worksheet.set_column(i, i, max_len)

        tactical_bytes.seek(0)

        # Save operational file
        operational_bytes = io.BytesIO()
        with pd.ExcelWriter(operational_bytes, engine='xlsxwriter') as writer:
            operational_output.to_excel(writer, index=False, sheet_name='Operational')
            workbook = writer.book
            worksheet = writer.sheets['Operational']
            date_format = workbook.add_format({'num_format': 'mm/dd/yyyy hh:mm'})

            for i, col in enumerate(operational_output.columns):
                max_len = max(operational_output[col].astype(str).map(len).max(), len(col)) + 2
                if col == "ENDEFF":
                    worksheet.set_column(i, i, max_len, date_format)
                else:
                    worksheet.set_column(i, i, max_len)

        operational_bytes.seek(0)

        st.download_button(
            label="📥 Download Tactical Output",
            data=tactical_bytes,
            file_name="tactical_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📥 Download Operational Output",
            data=operational_bytes,
            file_name="operational_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("🔍 Preview Tactical Output"):
            st.dataframe(tactical_output)

        with st.expander("🔍 Preview Operational Output"):
            st.dataframe(operational_output)

    except Exception as e:
        st.error(f"❌ Error occurred: {e}")
else:
    st.info("📂 Please upload both Excel files to begin.")

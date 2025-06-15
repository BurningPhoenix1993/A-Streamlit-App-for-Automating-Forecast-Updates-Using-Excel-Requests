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

# Only run if both files are uploaded
if data_manager_file and user_request_file:
    try:
        # Read files
        data_manager = pd.read_excel(data_manager_file, dtype=str).fillna('')
        user_request = pd.read_excel(user_request_file, dtype=str).fillna('')

        # Output containers
        tactical_rows = []
        operational_rows = []

        today_str = datetime.today().strftime("%m/%d/%Y")

        # Process each row
        for _, req in user_request.iterrows():
            filtered = data_manager.copy()

            # Apply filters
            if req.get("ProductId", "").strip():
                filtered = filtered[filtered["ProductId"] == req["ProductId"]]

            if req.get("SCHSALESCHANNELCD", "").strip():
                filtered = filtered[filtered["SCHSALESCHANNELCD"] == req["SCHSALESCHANNELCD"]]

            if req.get("LOCATIONID", "").strip():
                filtered = filtered[filtered["LOCATIONID"] == req["LOCATIONID"]]

            if req.get("IntegratedPLanningAccountId", "").strip():
                val = req["IntegratedPLanningAccountId"].strip()
                filtered = filtered[filtered["IntegratedPLanningAccountId"].str.startswith(val, na=False)]

            # Replace ProductId and ForecastItemId
            filtered = filtered.copy()
            filtered["ForecastItemId"] = filtered["ForecastItemId"].apply(
                lambda x: x.replace(req["ProductId"], req["New SKU"])
            )
            filtered["ProductId"] = req["New SKU"]

            # Set SOURCE and ENDEFF
            filtered["SOURCE"] = f"Manual Tarun Kumar {today_str}"
            filtered["ENDEFF"] = "12/31/2998 11:59:59 PM"

            # Final output columns
            output_cols = [
                "ForecastItemId",
                "ProductId",
                "CUSTOMERID",
                "LOCATIONID",
                "ForecastItemType",
                "SOURCE",
                "ENDEFF"
            ]

            # Separate into Tactical and Operational
            tactical_rows.append(filtered[filtered["ForecastItemType"] == "Tactical"][output_cols])
            operational_rows.append(filtered[filtered["ForecastItemType"] == "Operational"][output_cols])

        # Combine outputs
        tactical_output = pd.concat(tactical_rows, ignore_index=True)
        operational_output = pd.concat(operational_rows, ignore_index=True)

        st.success("✅ Processing Complete")

        # Save tactical file
        tactical_bytes = io.BytesIO()
        with pd.ExcelWriter(tactical_bytes, engine='xlsxwriter') as writer:
            tactical_output.to_excel(writer, index=False, sheet_name='Tactical')
            worksheet = writer.sheets['Tactical']
            for i, col in enumerate(tactical_output.columns):
                max_len = max(tactical_output[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        tactical_bytes.seek(0)

        # Save operational file
        operational_bytes = io.BytesIO()
        with pd.ExcelWriter(operational_bytes, engine='xlsxwriter') as writer:
            operational_output.to_excel(writer, index=False, sheet_name='Operational')
            worksheet = writer.sheets['Operational']
            for i, col in enumerate(operational_output.columns):
                max_len = max(operational_output[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        operational_bytes.seek(0)

        # Download buttons
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

        # Preview section
        with st.expander("🔍 Preview Tactical Output"):
            st.dataframe(tactical_output)

        with st.expander("🔍 Preview Operational Output"):
            st.dataframe(operational_output)

    except Exception as e:
        st.error(f"❌ Error occurred: {e}")
else:
    st.info("📂 Please upload both Excel files to begin.")

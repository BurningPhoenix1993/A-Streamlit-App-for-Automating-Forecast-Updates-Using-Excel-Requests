# A-Streamlit-App-for-Automating-Forecast-Updates-Using-Excel-Requests
This Streamlit-based web application automates the generation and update of Forecast Item IDs based on user-provided SKU mapping requests. It streamlines the process of updating tactical and operational forecast records using two uploaded Excel files: one containing master forecast data and another containing user modification requests.

This Streamlit app:
1.	Accepts Two Excel File Uploads:
o	One from a Data Manager source (with item/forecast data)
o	One with a User Request (indicating new SKU mappings and filters)
2.	Processes Each Request Row-by-Row:
o	Filters the data manager sheet using criteria like ProductId, SCHSALESCHANNELCD, LOCATIONID, and IntegratedPlanningAccountId
o	Replaces ProductId and ForecastItemId using the new SKU from the request
o	Adds metadata like SOURCE and a fixed ENDEFF (end effective date)
3.	Separates Output into Two Forecast Types:
o	Tactical
o	Operational
4.	Allows Downloading:
o	Generates downloadable Excel files for both tactical and operational outputs
o	Provides preview tables for both outputs within the app
5.	User-Friendly Interface:
o	Uses success/info/error messages, expanding previews, and automatic formatting in Excel

**Prerequisites**
pip install streamlit pandas openpyxl xlsxwriter (@command prompt)

**To run the app @command prompt**
streamlit run FI_creation.py


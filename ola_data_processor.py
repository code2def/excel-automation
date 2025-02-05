import streamlit as st
import pandas as pd

# Define user mapping
valid_users = {
    "abharti": "Ankur", "agiri1": "Aman", "dahuja": "Daksh", "dmam": "Deepak",
    "mranganathan": "Magesh", "psrihari": "Prakasam", "rjain6": "Rohit",
    "sarikapudi": "Sudheer", "sjain16": "Siddharth", "spatnam": "Sreekanth"
}

def process_excel(file):
    df = pd.read_excel(file, engine='pyxlsb')
    
    # Apply filters
    filtered_df = df[
        (df["QUEUE_CODE"] == "BDWCNFG") &
        (df["D_IN_OUT_OLA"] == "OUT OF OLA") &
        (df["USER_ID_COMPLETION"].isin(valid_users.keys()))
    ].copy()
    
    # Add Failure category & Failure Reasons
    if "Failure category" not in filtered_df.columns:
        filtered_df["Failure category"] = ""
    if "Failure Reasons" not in filtered_df.columns:
        filtered_df["Failure Reasons"] = ""
    
    filtered_df.loc[filtered_df["DELAY_DIARY"].notna(), "Failure category"] = "Genuine Fault / Prioritization Error"
    filtered_df.loc[filtered_df["DELAY_DIARY"].notna(), "Failure Reasons"] = filtered_df["USER_ID_COMPLETION"].map(
        lambda x: f"Missed to close on time by {valid_users.get(x, x)}"
    )
    
    # Define the correct column order
    correct_order = [
        "QUEUE_CODE", "TASK_CREATE", "TASK_CLOSED", "NEW_CONTRACT_NO", "ORDER_TYPE", "ORDER_TYPE_GROUP_CALC", "D_ORDER_TYPE_GROUP", "COUNTRY", "WORK_ITEM_ID_CALC", "REPORTING_WEEK", "REPORTING_MONTH", "PRODUCT_OFFERING", "D_OLA_TARGET", "LEAD_TIME_OVERALL", "D_IN_OUT_OLA", "REJECTION_DURATION_OVERALL", "CUSTOMER_ON_HOLD_DURITON_CALC", "TASK_DELAY_DURATION_OVERALL", "IN_OUT_OLA", "TASK_DELAY_REASON", "CUSTOMER_ON_HOLD_REASON", "NIGO_REASON_1", "NIGO_REASON_2", "SERVICE_GROUP_CALC", "TOP_250", "OPERATOR_NAME", "USER_ID_COMPLETION", "NO_OF_CIRCUITS", "INSTALLATION_COUNTRY", "CUSTOMER_NAME", "CUSTOMER_TIER", "NIGO_REASON", "INSTALLATION_COUNTRY_A", "INSTALLATION_COUNTRY_B", "A_COUNTRY", "B_COUNTRY", "SYS_FLAG", "D_CONNECTION_TYPE", "D_CONNECTION_TYPE_2", "CONTRACT_PRIORITY_ALL", "D_BANDWIDTH", "SD_WAN_SITE_TYPE", "FAST_TRACK", "FASTRACK_SIEBEL", "SUBTYPE", "DELAY_DIARY", "DELAY_REASON", "SIEBEL_PROJECT_ID", "SIEBEL_PROJECT_MANAGER", "OHS_PROJECT_ID", "OHS_PROJECT_NAME", "Location", "Sub Team", "Failure category", "Failure Reasons", "Comments", "Team Name", "Column2"
    ]
    
    # Reorder columns if all are present
    filtered_df = filtered_df.reindex(columns=correct_order, fill_value="")
    
    return filtered_df

# Streamlit UI
st.title("OLA Data Processor")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsb"])

if uploaded_file:
    st.write("Processing file...")
    result_df = process_excel(uploaded_file)
    
    st.write("### Filtered Data Preview:")
    st.dataframe(result_df)
    
    # Convert to downloadable Excel
    @st.cache
def convert_df(df):
        return df.to_csv(index=False).encode('utf-8')
    
    csv = convert_df(result_df)
    st.download_button("Download Processed File", csv, "filtered_data.csv", "text/csv")

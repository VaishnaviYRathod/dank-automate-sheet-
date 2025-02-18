import streamlit as st
import pandas as pd
import io

def find_header_row(df):
    """Find the first row with meaningful column names."""
    header_row_index = df.notna().sum(axis=1).idxmax()
    return header_row_index

def process_data(uploaded_file):
    try:
        # Read the uploaded Excel file
        raw_df = pd.read_excel(uploaded_file, None)  # Read all sheets
        sheet_name = list(raw_df.keys())[0]  # Use the first sheet
        df = raw_df[sheet_name]
        
        # Detect header row
        header_row_index = find_header_row(df)
        df = pd.read_excel(uploaded_file, skiprows=header_row_index)
        df.columns = df.iloc[0]  # Use detected row as header
        df = df[1:].reset_index(drop=True)  # Remove the header row from data
        
        # Define required columns
        required_columns = ['Sales Date *', 'POS Item ID *', 'POS Item Name', 'Total sales excl. tax *', 'Sold *']
        
        # Initialize processed DataFrame with required columns
        processed_df = pd.DataFrame(columns=required_columns)
        
        # Map matching columns from the uploaded file
        column_mapping = {
            'SR NO': 'POS Item ID *',
            'ITEMS': 'POS Item Name',
            'Total Amount': 'Total sales excl. tax *',
            'Sold': 'Sold *'
        }
        
        for col in df.columns:
            if col in column_mapping:
                processed_df[column_mapping[col]] = df[col]
        
        # Fill missing values with defaults
        processed_df['Sales Date *'] = pd.to_datetime('today').strftime('%d-%b-%y')
        processed_df.fillna({'POS Item ID *': 0, 'POS Item Name': 'Unknown', 'Total sales excl. tax *': 0, 'Sold *': 0}, inplace=True)
        
        return processed_df
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return None

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed Data')
    return output.getvalue()

# Streamlit UI
st.title("Automated Data Processing Platform")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])
if uploaded_file:
    st.write("Processing file...")
    processed_df = process_data(uploaded_file)
    
    if processed_df is not None:
        st.write("Processed Data Preview:")
        st.dataframe(processed_df.head())
        
        processed_file = convert_df_to_excel(processed_df)
        st.download_button(label="Download Processed File", data=processed_file, file_name="processed_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

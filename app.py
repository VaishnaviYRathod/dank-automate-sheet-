import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import numpy as np

st.image("supy_logo.png", width=200)

# Custom styling - Force all text to black, button/uploader text to light color

st.markdown(
    """
    <style>
    body {
        background-color: #f5f0ff;
        color: black !important; /* Force body text to black */
        font-family: Arial, sans-serif;
    }
    /* Force all labels to have black text */
    label {
        color: black !important;
    }
    /* Force all text within the app to black */
    .stApp * {
        color: black !important;
    }
    .stApp {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.1);
    }
    .stButton>button {
        background-color: #6a0dad;
        color: white !important; /* Keep button text white */
        font-weight: bold;
    }
    .stDownloadButton>button {
        background-color: #4b0082;
        color: white !important; /* Keep download button text white */
        font-weight: bold;
    }
    .stTitle {
        color: black !important;
    }
    .stFileUploader {
        border: 2px dashed #6a0dad;
        padding: 10px;
        background-color: #333333; /* Set a dark background for better contrast */
    }
    
    /* Make ALL text inside the uploader white */
    .stFileUploader label,
    .stFileUploader label div,
    .stFileUploader label span,
    .stFileUploader div,
    .stFileUploader span,
    .stFileUploader p,
    div[data-testid="stFileUploaderDropzone"] *,
    .stFileUploader div[data-testid="stFileUploaderDropzone"] * {
        color: white !important; /* Force all text to be white */
    }
    
    /* Specifically target each line to ensure they're white */
    /* "Drag and drop files here" text */
    .stFileUploader label > div:first-child {
        color: white !important;
    }
    
    /* "Limit 200MB per file" text */
    .stFileUploader label > div:nth-child(2) {
        color: white !important;
    }
    
    /* "• XLSX, XLS" text */
    .stFileUploader label > div:nth-child(3) {
        color: white !important;
    }
    
    .stFileUploader button {
        color: white !important; /* Ensure "Browse files" button text is white */
        background-color: #6a0dad; /* Match the primary button color */
        font-weight: bold;
    }
    .stCheckbox label span {
        color: black !important; /* Ensure checkbox labels are black */
    }
    .stRadio label span {
        color: black !important; /* Ensure radio button labels are black */
    }
    .stSelectbox label span {
        color: black !important; /* Ensure selectbox labels are black */
    }
    .stMultiselect label span {
        color: black !important; /* Ensure multiselect labels are black */
    }
    .stSlider label span {
        color: black !important; /* Ensure slider labels are black */
    }
    .stNumberInput label span {
        color: black !important; /* Ensure number input labels are black */
    }
    .stTextArea label span {
        color: black !important; /* Ensure textarea labels are black */
    }
    .stTextInput label span {
        color: black !important; /* Ensure text input labels are black */
    }
    .stMarkdown {
        color: black !important; /* Ensure markdown text is black */
    }
    .stTable {
        color: black !important; /* Ensure table text is black */
    }
    .stDataFrame {
        color: black !important; /* Ensure dataframe text is black */
    }
    .stSuccess {
        color: black !important; /* Ensure success messages are black */
    }
    .stError {
        color: black !important; /* Ensure error messages are black */
    }
    .stWarning {
        color: black !important; /* Ensure warning messages are black */
    }
    .stInfo {
        color: black !important; /* Ensure info messages are black */
    }
    </style>
    """,
    unsafe_allow_html=True
)
# Global variable for debug info
show_debug_info = False

# Improved function to extract the sales date from the sheet name (now from Excel sheet name)
def extract_date_from_sheet_name(sheet_name):
    # Normalize sheet name by trimming whitespace
    sheet_name = sheet_name.strip()

    # Additional date formats to try
    date_formats = [
        "%B %d, %Y",      # March 1, 2025
        "%b %d, %Y",      # Mar 1, 2025
        "%B %d,%Y",       # March 1,2025
        "%b %d,%Y",       # Mar 1,2025
        "%d %B %Y",       # 1 March 2025
        "%d %b %Y",       # 1 Mar 2025
        "%d-%B-%Y",       # 1-March-2025
        "%d-%b-%Y",       # 1-Mar-2025
        "%B-%d-%Y",       # March-1-2025
        "%b-%d-%Y",       # Mar-1-2025
        "%m/%d/%Y",       # 3/1/2025
        "%d/%m/%Y",       # 1/3/2025
        "%Y-%m-%d"        # 2025-03-01
    ]

    # Try parsing with each format
    for fmt in date_formats:
        try:
            return datetime.strptime(sheet_name, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    # If direct parsing failed, use regex to extract parts
    patterns = [
        # Month spelled out, then day, then year
        r'([A-Za-z]+)\s+(\d{1,2})(?:[,\s]+)(\d{4})',
        # Day, then month, then year
        r'(\d{1,2})\s+([A-Za-z]+)(?:[,\s]+)(\d{4})',
        # Various numeric formats
        r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})',
        r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})'
    ]

    for pattern in patterns:
        match = re.search(pattern, sheet_name)
        if match:
            groups = match.groups()
            try:
                if len(groups[0]) > 2:  # First group is month name
                    month, day, year = groups
                    date_str = f"{month} {day}, {year}"
                    for fmt in ["%B %d, %Y", "%b %d, %Y"]:
                        try:
                            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                        except ValueError:
                            continue
                elif groups[0].isdigit() and len(groups[0]) == 4:  # First group is year
                    year, month, day = groups
                    return f"{year}-{int(month):02d}-{int(day):02d}"
                else:  # First group is day or month number
                    if len(groups[1]) > 2:  # Second group is month name
                        day, month, year = groups
                        date_str = f"{day} {month} {year}"
                        for fmt in ["%d %B %Y", "%d %b %Y"]:
                            try:
                                return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
                            except ValueError:
                                continue
                    else:  # Numeric date format
                        g1, g2, g3 = groups
                        # Try both MM/DD/YYYY and DD/MM/YYYY interpretations
                        try:
                            return datetime.strptime(f"{g1}/{g2}/{g3}", "%m/%d/%Y").strftime("%Y-%m-%d")
                        except ValueError:
                            try:
                                return datetime.strptime(f"{g1}/{g2}/{g3}", "%d/%m/%Y").strftime("%Y-%m-%d")
                            except ValueError:
                                pass
            except Exception:
                pass

    # Special case for sheets with "March" in the name but without day/year
    month_pattern = re.compile(r'\b(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b', re.IGNORECASE)
    month_match = month_pattern.search(sheet_name)

    if month_match:
        # If only month is found, use current year and day 1 as defaults
        month_name = month_match.group(1)
        try:
            # Try with full month name
            date_str = f"{month_name} 1, 2025"
            return datetime.strptime(date_str, "%B 1, %Y").strftime("%Y-%m-%d")
        except ValueError:
            try:
                # Try with abbreviated month name
                date_str = f"{month_name} 1, 2025"
                return datetime.strptime(date_str, "%b 1, %Y").strftime("%Y-%m-%d")
            except ValueError:
                pass

    # If we can't extract a date, return None but don't show a warning for known non-date sheet names
    if sheet_name.lower() in ["sales report", "daily sales report", "summary", "data"]:
        return None

    # Log the sheet name that couldn't be parsed
    if show_debug_info:
        st.warning(f"Could not extract date from sheet name: '{sheet_name}'")
    return None

# Function to detect the header row
def detect_header_row(df):
    header_keywords = ["sr no", "sr. no", "items", "beginning", "receival", "sold", "write-off", "end count", "variance", "unit price", "total amount", "expiry date"]
    for i in range(min(15, len(df))):  # Check first 15 rows for headers (increased from 10)
        row_values = df.iloc[i].astype(str).str.lower()
        matches = sum(any(keyword in str(cell).lower() for keyword in header_keywords) for cell in row_values)
        if matches >= 2:  # If at least 2 header keywords are found
            return i
    return 0

# Function to clean and process a single sheet
def process_sales_data(sheet_name, df):
    try:
        # Skip empty sheets
        if df.empty:
            if show_debug_info:
                st.warning(f"Sheet '{sheet_name}' is empty, skipping.")
            return pd.DataFrame()

        # Detect header row
        header_row = detect_header_row(df)

        # If DataFrame doesn't have enough rows, return empty DataFrame
        if len(df) <= header_row + 1:
            if show_debug_info:
                st.warning(f"Sheet '{sheet_name}' doesn't have enough data rows after header row {header_row}.")
            return pd.DataFrame()

        # Set header and filter data rows
        df.columns = df.iloc[header_row]
        df = df[header_row + 1:].reset_index(drop=True)

        # Clean column names
        df = df.rename(columns=lambda x: str(x).strip())

        # Save a copy of the original column names (both upper and lower case for later matching)
        original_columns = {col.upper(): col for col in df.columns}

        # Convert column names to uppercase for standardized processing
        df.columns = df.columns.str.upper()

        # Display column names for debugging
        if show_debug_info:
            st.write(f"Columns in sheet '{sheet_name}' after header detection: {list(df.columns)}")

        # Find the first "TOTAL" row index to only process data until that point
        total_row_idx = None
        for idx, row in df.iterrows():
            row_values = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("TOTAL" in x for x in row_values):
                total_row_idx = idx
                break

        # If a TOTAL row is found, only process data up to that row
        if total_row_idx is not None:
            if show_debug_info:
                st.info(f"Found TOTAL row at index {total_row_idx} in sheet '{sheet_name}'")
            df = df.iloc[:total_row_idx]

        # Extract sales date from sheet name
        sales_date = extract_date_from_sheet_name(sheet_name)

        # Log the extracted date for debugging
        if sales_date and show_debug_info:
            st.info(f"Extracted date '{sales_date}' from sheet name: '{sheet_name}'")

        df["SALES DATE *"] = sales_date if sales_date else ""

        # Define column mapping based on your requirements
        column_mapping = {}

        # Look for SR NO column for POS Item ID
        sr_no_cols = [col for col in df.columns if 'SR' in col and ('NO' in col or 'NUMBER' in col)]
        if sr_no_cols:
            column_mapping[sr_no_cols[0]] = "POS ITEM ID *"

        # Look for ITEMS column for POS Item Name
        item_cols = [col for col in df.columns if any(x in col for x in ['ITEM', 'ITEMS', 'PRODUCT', 'DESCRIPTION'])]
        if item_cols:
            column_mapping[item_cols[0]] = "POS ITEM NAME"

        # Look for TOTAL AMOUNT column for sales values
        total_amount_cols = [col for col in df.columns if
                                    ('TOTAL' in col and 'AMOUNT' in col) or
                                    ('TOTAL' in col and 'SALES' in col) or
                                    ('TOTAL' in col and 'PRICE' in col) or
                                    ('TOTAL' in col and 'COST' in col)]
        if total_amount_cols:
            column_mapping[total_amount_cols[0]] = "TOTAL SALES INCL. TAX *"

        # Look for SOLD column for Sold QTY
        sold_cols = [col for col in df.columns if any(x in col for x in ['SOLD', 'SALE QTY', 'QTY SOLD', 'QUANTITY'])]
        if sold_cols:
            column_mapping[sold_cols[0]] = "SOLD QTY *"

        # Display column mapping for debugging
        if show_debug_info:
            st.write(f"Column mapping for sheet '{sheet_name}': {column_mapping}")

        # Rename columns based on mapping
        df.rename(columns=column_mapping, inplace=True)

        # Ensure "TOTAL SALES INCL. TAX *" is present
        if "TOTAL SALES INCL. TAX *" not in df.columns:
            # Look for any price/amount/total columns
            price_cols = [col for col in df.columns if any(term in col
                                                            for term in ["PRICE", "AMOUNT", "COST", "TOTAL", "SALE"])]
            if price_cols:
                df.rename(columns={price_cols[0]: "TOTAL SALES INCL. TAX *"}, inplace=True)
            else:
                df["TOTAL SALES INCL. TAX *"] = 0.0

        # Make TOTAL SALES EXCL. TAX * equal to TOTAL SALES INCL. TAX * as requested
        df["TOTAL SALES EXCL. TAX *"] = df["TOTAL SALES INCL. TAX *"]

        # Ensure POS ITEM ID column exists
        if "POS ITEM ID *" not in df.columns:
            # Use index as item ID if not present
            df["POS ITEM ID *"] = df.index + 1

        # Ensure POS ITEM NAME column exists
        if "POS ITEM NAME" not in df.columns:
            # Try to find a column that might contain item names
            possible_name_cols = [col for col in df.columns if any(term in col
                                                                    for term in ["ITEM", "PRODUCT", "DESCRIPTION"])]
            if possible_name_cols:
                df.rename(columns={possible_name_cols[0]: "POS ITEM NAME"}, inplace=True)
            else:
                df["POS ITEM NAME"] = f"Item from {sheet_name}"

        # Ensure SOLD QTY column exists
        if "SOLD QTY *" not in df.columns:
            # Try to find quantity-related columns, with more variations
            qty_cols = [col for col in df.columns if any(term in col
                                                            for term in ["QTY", "QUANTITY", "SOLD", "COUNT", "SALE"])]
            if qty_cols:
                df.rename(columns={qty_cols[0]: "SOLD QTY *"}, inplace=True)
            else:
                df["SOLD QTY *"] = 1  # Default to 1

        # Define required columns (now in uppercase to match our processed columns)
        required_columns = ["SALES DATE *", "POS ITEM ID *", "POS ITEM NAME", "SOLD QTY *",
                            "TOTAL DISCOUNT VALUE", "TOTAL SALES EXCL. TAX *", "TOTAL SALES INCL. TAX *",
                            "ORDER ID", "SALES TYPE CODE"]

        # Ensure all required columns are present
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""  # Fill missing columns with empty values

        # Clean and process numeric columns
        numeric_columns = ["SOLD QTY *", "TOTAL SALES EXCL. TAX *", "TOTAL SALES INCL. TAX *", "TOTAL DISCOUNT VALUE"]
        for col in numeric_columns:
            if col in df.columns:
                # Convert to string for consistent preprocessing
                df[col] = df[col].astype(str)
                # Replace commas and other non-numeric characters
                df[col] = df[col].str.replace(',', '').str.replace('$', '').str.replace('₹', '').str.strip()
                # Convert to numeric values
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Filter out rows with invalid or zero quantities for SOLD QTY
        if "SOLD QTY *" in df.columns:
            df = df[df["SOLD QTY *"] > 0]
            # Convert to integer
            df["SOLD QTY *"] = df["SOLD QTY *"].astype(int)

        # Create final processed DataFrame with only required columns
        processed_df = pd.DataFrame()
        for col in required_columns:
            processed_df[col] = df[col] if col in df.columns else ""

        # Remove any leftover 'TOTAL' rows (as a safeguard)
        if "POS ITEM NAME" in processed_df.columns:
            processed_df = processed_df[~processed_df["POS ITEM NAME"].astype(str).str.contains("TOTAL", case=False, na=False)]

        # Title case the column names for final output (to match expected format)
        processed_df.columns = [col.title() for col in processed_df.columns]

        return processed_df

    except Exception as e:
        if show_debug_info:
            st.error(f"Error processing sheet '{sheet_name}': {str(e)}")
        return pd.DataFrame()  # Return empty DataFrame on error

def main():
    st.title("Bulk Excel Sales Data Processor")

    st.info("""
    ### Instructions:
    1. Upload Excel files (.xlsx or .xls). You can upload multiple files.
    2. Click "Process Files".
    3. Review the processed data for each sheet in each file (up to the first 5 files uploaded).
    4. Download the combined processed data as a single Excel file.
    """)

    uploaded_files = st.file_uploader("Upload Excel Files", type=["xlsx", "xls"], accept_multiple_files=True)

    all_processed_data = []

    if uploaded_files:
        if len(uploaded_files) > 5:
            st.warning("You have uploaded more than 5 files. Only the first 5 will be processed.")
            files_to_process = uploaded_files[:5]
        else:
            files_to_process = uploaded_files

        for uploaded_file in files_to_process:
            st.subheader(f"Processing File: {uploaded_file.name}")
            try:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names

                for sheet_name in sheet_names:
                    st.subheader(f"- Processing Sheet: {sheet_name} from {uploaded_file.name}")
                    df = excel_file.parse(sheet_name)
                    if not df.empty:
                        processed_df = process_sales_data(sheet_name, df.copy())
                        if not processed_df.empty:
                            all_processed_data.append(processed_df)
                            st.dataframe(processed_df)
                        else:
                            st.warning(f"-- No valid data found after processing sheet: {sheet_name} from {uploaded_file.name}")
                    else:
                        st.warning(f"-- Sheet '{sheet_name}' in {uploaded_file.name} is empty.")

            except Exception as e:
                st.error(f"An error occurred while processing file '{uploaded_file.name}': {e}")

        if all_processed_data:
            combined_df = pd.concat(all_processed_data, ignore_index=True)
            st.subheader("Combined Processed Data from Processed Files")
            st.dataframe(combined_df)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, sheet_name='Processed Data', index=False)
            output.seek(0)

            st.download_button(
                label="Download Combined Processed Data (Excel)",
                data=output.getvalue(),
                file_name="combined_processed_sales_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No data was processed from the uploaded files.")

if __name__ == "__main__":
    main()

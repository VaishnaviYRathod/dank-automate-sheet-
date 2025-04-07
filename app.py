import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import numpy as np

st.image("supy_logo.png", width=200)

# Custom styling
st.markdown(
    """
    <style>
    body {
        background-color: #f5f0ff;
        color: #4b0082;
        font-family: Arial, sans-serif;
    }
    <style>
/* Force all labels to have black text */
label {
    color: black !important;
}
</style>

    .stApp {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.1);
    }
    .stButton>button {
        background-color: #6a0dad;
        color: white;
        font-weight: bold;
    }
    .stDownloadButton>button {
        background-color: #4b0082;
        color: white;
        font-weight: bold;
    }
    .stTitle {
        color: #4b0082;
    }
    .stFileUploader {
        border: 2px dashed #6a0dad;
        padding: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Global variable for debug info
show_debug_info = False

# Improved function to extract the sales date from the sheet name
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

# Streamlit UI
st.markdown(
    """
    <h1 style='color: black;'>Sales Data Processor</h1>
    <p style='color: black; font-size: 18px;'>Upload your sales data file and get a cleaned, structured version.</p>
    """,
    unsafe_allow_html=True
)

# Add checkbox for displaying debug information
show_debug_info = st.checkbox("Show debug information", value=False)

# Add option to include total row (off by default)
include_total_row = st.checkbox("Include total row in output", value=False)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    st.write("Processing...")
    
    try:
        # Display sheet names for debugging
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        if show_debug_info:
            st.write("Sheet names found in the file:")
            for i, name in enumerate(sheet_names):
                date = extract_date_from_sheet_name(name)
                if date:
                    st.write(f"- {name} ✓ (Date extracted: {date})")
                else:
                    st.write(f"- {name} ✗ (No date extracted)")
        
        # Read all sheets
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        processed_sheets = []
        successful_sheets = 0
        failed_sheets = 0
        
        # Process each sheet
        for sheet_name, df in all_sheets.items():
            if show_debug_info:
                st.write(f"Processing sheet: '{sheet_name}'")
            
            processed_df = process_sales_data(sheet_name, df)
            
            if not processed_df.empty:
                processed_sheets.append(processed_df)
                successful_sheets += 1
            else:
                failed_sheets += 1
        
        # Combine all processed sheets
        if processed_sheets:
            final_data = pd.concat(processed_sheets, ignore_index=True)
            
            if not final_data.empty:
                # Only add total row if requested
                if include_total_row:
                    # Compute total row
                    total_row = pd.Series(dtype='object')
                    for col in final_data.columns:
                        if col == "Sales Date *":
                            total_row[col] = "Total"
                        elif col in ["Pos Item Name", "Order Id", "Sales Type Code"]:
                            total_row[col] = ""
                        else:
                            # Try to sum numeric columns
                            try:
                                total_row[col] = pd.to_numeric(final_data[col], errors='coerce').sum()
                            except:
                                total_row[col] = ""
                    
                    final_data = pd.concat([final_data, pd.DataFrame([total_row])], ignore_index=True)
                
                # Display processing summary
                st.success(f"File processed successfully! {successful_sheets} sheets processed, {failed_sheets} sheets skipped.")
                
                # Show sample of processed data
                st.write("Preview of processed data:")
                st.dataframe(final_data.head())
                
                # Convert to Excel and create download button
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_data.to_excel(writer, index=False, sheet_name="Processed Data")
                
                st.download_button(
                    label="Download Processed File",
                    data=output.getvalue(),
                    file_name="Processed_Sales_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No valid data found in the processed sheets.")
        else:
            st.error(f"No valid data found in the uploaded file. All {failed_sheets} sheets were skipped.")
    
    except Exception as e:
        st.error(f"An error occurred during processing: {str(e)}")

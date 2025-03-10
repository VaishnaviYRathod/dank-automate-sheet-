import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import numpy as np
import os

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
# Improved function to extract the sales date from the sheet name
# Fixed function to extract the sales date from the sheet name
def extract_date_from_sheet_name(sheet_name):
    # Normalize sheet name by trimming whitespace
    sheet_name = sheet_name.strip()
    
    # Log the sheet name we're trying to extract date from
    if show_debug_info:
        st.write(f"Attempting to extract date from sheet name: '{sheet_name}'")
    
    # Try direct parsing first - these are the exact formats we're looking for
    exact_formats = [
        "%b %d, %Y",      # Feb 1, 2025
        "%B %d, %Y",      # February 1, 2025
        "%b %d,%Y",       # Feb 1,2025
        "%B %d,%Y"        # February 1,2025
    ]
    
    for fmt in exact_formats:
        try:
            parsed_date = datetime.strptime(sheet_name, fmt).strftime("%Y-%m-%d")
            if show_debug_info:
                st.success(f"Successfully parsed '{sheet_name}' using format '{fmt}' to '{parsed_date}'")
            return parsed_date
        except ValueError:
            continue
    
    # If direct parsing fails, use regex to extract parts
    month_abbr_pattern = re.compile(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),?\s+(\d{4})\b', re.IGNORECASE)
    month_full_pattern = re.compile(r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b', re.IGNORECASE)
    
    # Month name mapping
    month_map = {
        'jan': 1, 'january': 1, 
        'feb': 2, 'february': 2, 
        'mar': 3, 'march': 3, 
        'apr': 4, 'april': 4, 
        'may': 5, 'may': 5,
        'jun': 6, 'june': 6, 
        'jul': 7, 'july': 7, 
        'aug': 8, 'august': 8, 
        'sep': 9, 'september': 9, 
        'oct': 10, 'october': 10, 
        'nov': 11, 'november': 11, 
        'dec': 12, 'december': 12
    }
    
    # Check for abbreviated month names first (Feb 1, 2025)
    match = month_abbr_pattern.search(sheet_name)
    if match:
        month_str, day, year = match.groups()
        month_num = month_map.get(month_str.lower(), 1)
        return f"{year}-{month_num:02d}-{int(day):02d}"
    
    # Then check for full month names (February 1, 2025)
    match = month_full_pattern.search(sheet_name)
    if match:
        month_str, day, year = match.groups()
        month_num = month_map.get(month_str.lower(), 1)
        return f"{year}-{month_num:02d}-{int(day):02d}"
    
    # If specific formats fail, try broader patterns
    # Additional date formats to try
    date_formats = [
        "%m/%d/%Y",       # 2/1/2025
        "%d/%m/%Y",       # 1/2/2025
        "%Y-%m-%d",       # 2025-02-01
        "%m-%d-%Y",       # 02-01-2025
        "%d-%m-%Y",       # 01-02-2025
    ]
    
    # Try parsing with additional formats
    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(sheet_name, fmt).strftime("%Y-%m-%d")
            if show_debug_info:
                st.success(f"Successfully parsed '{sheet_name}' using format '{fmt}' to '{parsed_date}'")
            return parsed_date
        except ValueError:
            continue
    
    # Fallback to a default date instead of None
    # This ensures the Sales Date column always has a value
    default_date = "2025-01-01"
    if show_debug_info:
        st.warning(f"Could not extract date from sheet name: '{sheet_name}'. Using default date: {default_date}")
    
    return default_date  # Default to January 1, 2025 for sheets without dates

    

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

# Function to extract date from filename
def extract_date_from_filename(filename):
    # Remove file extension
    base_name = os.path.splitext(filename)[0]
    
    # Try to extract date using the same techniques as for sheet names
    return extract_date_from_sheet_name(base_name)

# Function to detect the header row
def detect_header_row(df):
    header_keywords = ["sr no", "sr. no", "items", "beginning", "receival", "sold", "write-off", "end count", "variance", "unit price", "total amount", "expiry date"]
    for i in range(min(15, len(df))):  # Check first 15 rows for headers
        row_values = df.iloc[i].astype(str).str.lower()
        matches = sum(any(keyword in str(cell).lower() for keyword in header_keywords) for cell in row_values)
        if matches >= 2:  # If at least 2 header keywords are found
            return i
    return 0

# Function to clean and process a single sheet
def process_sales_data(sheet_name, df, file_date=None):
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
        
        # Extract sales date from sheet name or use provided file date
        sales_date = extract_date_from_sheet_name(sheet_name)
        
        # If sheet name doesn't have a date, use file date
        if not sales_date and file_date:
            sales_date = file_date
            if show_debug_info:
                st.info(f"Using file date '{sales_date}' for sheet '{sheet_name}'")
        
        # Log the extracted date for debugging
        if sales_date and show_debug_info:
            st.info(f"Using date '{sales_date}' for sheet '{sheet_name}'")
        
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
        
        # Create a new DataFrame with only the required columns
        new_df = pd.DataFrame()
        
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
        required_columns = ["POS ITEM ID *", "POS ITEM NAME", "SOLD QTY *", 
                           "TOTAL DISCOUNT VALUE", "TOTAL SALES EXCL. TAX *", "TOTAL SALES INCL. TAX *", 
                           "ORDER ID", "SALES TYPE CODE"]
        
        # First, add the sales date
        new_df["SALES DATE *"] = sales_date if sales_date else ""
        
        # Copy only the required columns from df to new_df
        for col in required_columns:
            if col in df.columns:
                new_df[col] = df[col]
            else:
                new_df[col] = ""  # Fill missing columns with empty values
        
        # Clean and process numeric columns
        numeric_columns = ["SOLD QTY *", "TOTAL SALES EXCL. TAX *", "TOTAL SALES INCL. TAX *", "TOTAL DISCOUNT VALUE"]
        for col in numeric_columns:
            if col in new_df.columns:
                # Convert to string for consistent preprocessing
                new_df[col] = new_df[col].astype(str)
                # Replace commas and other non-numeric characters
                new_df[col] = new_df[col].str.replace(',', '').str.replace('$', '').str.replace('₹', '').str.strip()
                # Convert to numeric values
                new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0)
        
        # Filter out rows with invalid or zero quantities for SOLD QTY
        if "SOLD QTY *" in new_df.columns:
            new_df = new_df[new_df["SOLD QTY *"] > 0]
            # Convert to integer
            new_df["SOLD QTY *"] = new_df["SOLD QTY *"].astype(int)
        
        # Remove any leftover 'TOTAL' rows (as a safeguard)
        if "POS ITEM NAME" in new_df.columns:
            new_df = new_df[~new_df["POS ITEM NAME"].astype(str).str.contains("TOTAL", case=False, na=False)]
        
        # Title case the column names for final output (to match expected format)
        new_df.columns = [col.title() for col in new_df.columns]
        
        return new_df
    
    except Exception as e:
        if show_debug_info:
            st.error(f"Error processing sheet '{sheet_name}': {str(e)}")
        return pd.DataFrame()  # Return empty DataFrame on error

# Function to process a single file
def process_file(uploaded_file):
    try:
        # Extract date from filename if possible
        file_date = extract_date_from_filename(uploaded_file.name)
        if file_date and show_debug_info:
            st.info(f"Extracted date '{file_date}' from filename: '{uploaded_file.name}'")
        
        # Get sheet names
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        if show_debug_info:
            st.write(f"Processing file: {uploaded_file.name}")
            st.write("Sheet names found in the file:")
            for i, name in enumerate(sheet_names):
                sheet_date = extract_date_from_sheet_name(name)
                date_message = f"(Date extracted: {sheet_date})" if sheet_date else "(No date extracted, using file date if available)"
                st.write(f"- {name} {date_message}")
        
        # Read all sheets
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        processed_sheets = []
        successful_sheets = 0
        failed_sheets = 0
        
        # Process each sheet
        for sheet_name, df in all_sheets.items():
            if show_debug_info:
                st.write(f"Processing sheet: '{sheet_name}'")
            
            processed_df = process_sales_data(sheet_name, df, file_date)
            
            if not processed_df.empty:
                processed_sheets.append(processed_df)
                successful_sheets += 1
            else:
                failed_sheets += 1
        
        # Combine all processed sheets from this file
        if processed_sheets:
            file_data = pd.concat(processed_sheets, ignore_index=True)
            if show_debug_info:
                st.success(f"File '{uploaded_file.name}' processed: {successful_sheets} sheets processed, {failed_sheets} sheets skipped.")
            return file_data
        else:
            if show_debug_info:
                st.error(f"No valid data found in file '{uploaded_file.name}'. All {failed_sheets} sheets were skipped.")
            return pd.DataFrame()
    
    except Exception as e:
        if show_debug_info:
            st.error(f"Error processing file '{uploaded_file.name}': {str(e)}")
        return pd.DataFrame()

# Streamlit UI
st.markdown(
    """
    <h1 style='color: black;'>Multi-File Sales Data Processor</h1>
    <p style='color: black; font-size: 18px;'>Upload up to 5 sales data files and get a cleaned, combined dataset.</p>
    """,
    unsafe_allow_html=True
)

# Add checkbox for displaying debug information
show_debug_info = st.checkbox("Show debug information", value=False)

# Allow multiple file uploads (up to 5)
uploaded_files = st.file_uploader("Upload Excel Files (up to 5)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # Check if we have more than 5 files
    if len(uploaded_files) > 5:
        st.warning(f"You've uploaded {len(uploaded_files)} files. Only the first 5 will be processed.")
        uploaded_files = uploaded_files[:5]
    
    st.write(f"Processing {len(uploaded_files)} files...")
    
    try:
        all_processed_data = []
        successful_files = 0
        failed_files = 0
        
        # Process each file
        progress_bar = st.progress(0)
        for i, file in enumerate(uploaded_files):
            file_data = process_file(file)
            
            if not file_data.empty:
                all_processed_data.append(file_data)
                successful_files += 1
            else:
                failed_files += 1
            
            # Update progress bar
            progress_bar.progress((i + 1) / len(uploaded_files))
        
        progress_bar.empty()
        
        # Combine all processed data
        if all_processed_data:
            combined_data = pd.concat(all_processed_data, ignore_index=True)
            
            if not combined_data.empty:
                # Define the exact required columns in the desired order
                required_output_columns = [
                    "Sales Date *", "Pos Item Id *", "Pos Item Name", "Sold Qty *", 
                    "Total Discount Value", "Total Sales Excl. Tax *", "Total Sales Incl. Tax *", 
                    "Order Id", "Sales Type Code"
                ]
                
                # Create a new DataFrame with only the required columns in the correct order
                final_output = pd.DataFrame()
                for col in required_output_columns:
                    if col in combined_data.columns:
                        final_output[col] = combined_data[col]
                    else:
                        final_output[col] = ""
                
                # Compute total row
                total_row = pd.Series(dtype='object')
                for col in required_output_columns:
                    if col == "Sales Date *":
                        total_row[col] = "Total"
                    elif col in ["Pos Item Name", "Order Id", "Sales Type Code", "Pos Item Id *"]:
                        total_row[col] = ""
                    else:
                        # Try to sum numeric columns
                        try:
                            total_row[col] = pd.to_numeric(final_output[col], errors='coerce').sum()
                        except:
                            total_row[col] = ""
                
                final_output = pd.concat([final_output, pd.DataFrame([total_row])], ignore_index=True)
                
                # Display processing summary
                st.success(f"Processing complete! {successful_files} files processed successfully, {failed_files} files failed.")
                
                # Show sample of processed data with explicit date display
                st.write("Preview of processed data:")
                
                # Format the Sales Date column for better display
                display_df = final_output.head().copy()
                if "Sales Date *" in display_df.columns:
                    # Convert empty strings to explicit "Not Available" for clear display
                    display_df["Sales Date *"] = display_df["Sales Date *"].apply(lambda x: x if x else "Not Available")
                
                st.dataframe(display_df)
                
                # Show stats
                total_items = len(final_output) - 1  # Subtract 1 for the total row
                total_sales = final_output.iloc[-1]["Total Sales Incl. Tax *"] if "Total Sales Incl. Tax *" in final_output.columns else 0
                
                st.info(f"Total Items: {total_items}, Total Sales: ${total_sales:.2f}")
                
                # Convert to Excel and create download button
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_output.to_excel(writer, index=False, sheet_name="Processed Data")
                
                st.download_button(
                    label="Download Combined Processed File",
                    data=output.getvalue(),
                    file_name="Combined_Sales_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No valid data found in any of the processed files.")
        else:
            st.error(f"No valid data found in any of the uploaded files. All {failed_files} files were skipped.")
    
    except Exception as e:
        st.error(f"An error occurred during processing: {str(e)}")

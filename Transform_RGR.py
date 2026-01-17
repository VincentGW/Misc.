import xlwings as xw
import pandas as pd
import os
import traceback
import sys
import json

# Set pandas display options to prevent truncation
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Define paths - use the directory where this script is located
base_path = os.path.dirname(os.path.abspath(__file__))

# Find the Excel file that starts with "RGR"
import glob
rgr_files = glob.glob(os.path.join(base_path, "RGR*.xlsx"))

if not rgr_files:
    print("ERROR: No Excel file starting with 'RGR' found in the directory.")
    input("\nPress Enter to exit...")
    exit()
elif len(rgr_files) > 1:
    print(f"WARNING: Multiple files starting with 'RGR' found. Using the first one: {os.path.basename(rgr_files[0])}")
    start_file = rgr_files[0]
else:
    start_file = rgr_files[0]
    print(f"Found input file: {os.path.basename(start_file)}")

list_file = os.path.join(base_path, "List.xlsx")
rates_file = os.path.join(base_path, "Online Tuition Rates.txt")

print("Reading Excel Files...\n")

try:
    with open(rates_file, 'r') as f:
        tuition_rates = json.load(f)
        undergraduate_rate = tuition_rates["UGRD"]
        graduate_rate = tuition_rates["GRAD"]

    # Silently open Start.xlsx
    app = xw.App(visible=False)
    wb = xw.Book(start_file)

    # Get the first sheet
    sheet = wb.sheets[0]

    # Read everything beneath the first row into a dataframe
    # Skip the first row (skiprows=1) and use the second row as header
    data = sheet.range("A2").expand().options(pd.DataFrame, header=1, index=False).value

    # Find the first occurrence of Career and Campus
    cols = list(data.columns)
    career_indices = [i for i, col in enumerate(cols) if col == "Career"]
    campus_indices = [i for i, col in enumerate(cols) if col == "Campus"]

    # Rename duplicate column names to make them unique
    new_cols = []
    col_counts = {}
    for col in cols:
        if col in col_counts:
            col_counts[col] += 1
            new_cols.append(f"{col}_{col_counts[col]}")
        else:
            col_counts[col] = 0
            new_cols.append(col)

    data.columns = new_cols

    # Create a copy of data for processing
    process = data.copy()

    # Now select only the columns we want (using renamed column names)
    columns_to_keep = []
    for col in process.columns:
        if col in ["ID", "Last", "First Name", "Career", "Campus"]:
            columns_to_keep.append(col)

    # Keep only the selected columns in process
    process = process[columns_to_keep]

    # Remove duplicates by ID and Career (keep only unique combinations)
    process = process.drop_duplicates(subset=["ID", "Career"], keep="first").reset_index(drop=True)

    # Load List.xlsx as dataframe gs
    list_wb = xw.Book(list_file)
    list_sheet = list_wb.sheets[0]
    gs = list_sheet.range("A1").expand().options(pd.DataFrame, header=1, index=False).value

    # Rename UID to ID in gs
    if 'UID' in gs.columns:
        gs.rename(columns={'UID': 'ID'}, inplace=True)

    # Convert both columns to strings (without decimals)
    # First column is Term, second column is ID
    # Create new columns to avoid dtype incompatibility warnings
    term_col_name = gs.columns[0]
    id_col_name = gs.columns[1]

    gs[term_col_name] = pd.to_numeric(gs[term_col_name], errors='coerce').fillna(0).astype(int).astype(str)
    gs[id_col_name] = pd.to_numeric(gs[id_col_name], errors='coerce').fillna(0).astype(int).astype(str)

    # Close List workbook
    list_wb.close()

    def get_terms_for_id(uid):
        # Filter gs by ID column
        filtered = gs[gs['ID'] == uid]
        # Get the Term column (first column) as a list
        terms = filtered.iloc[:, 0].tolist()
        # Return up to 3 terms, pad with empty strings if needed
        py1 = terms[0] if len(terms) > 0 else ""
        py2 = terms[1] if len(terms) > 1 else ""
        py3 = terms[2] if len(terms) > 2 else ""
        return py1, py2, py3

    # Apply the function to create the three columns in process
    process[['py1', 'py2', 'py3']] = process['ID'].apply(
        lambda x: pd.Series(get_terms_for_id(x))
    )

    # Look for columns that contain "Term" in the name
    term_columns = [col for col in data.columns if 'Term' in col]

    # Collect all unique terms from these columns
    all_terms = set()
    for col in term_columns:
        # Get unique values, convert to string, remove decimals
        unique_vals = data[col].dropna().unique()
        for val in unique_vals:
            # Convert to int then string to remove decimals
            try:
                term_str = str(int(float(val)))
                all_terms.add(term_str)
            except:
                pass

    # Sort terms from smallest to largest
    sorted_terms = sorted(all_terms, key=lambda x: int(x))
    print(f"Found {len(sorted_terms)} unique terms")
    print(f"Terms (sorted): {sorted_terms}")
    print()

    # Create columns in process for each unique term (initialized with empty strings)
    for term in sorted_terms:
        process[term] = ""

    # First, find the Unit Taken column in data
    units_columns = [col for col in data.columns if 'Unit Taken' in col or 'Units Taken' in col or 'Units' in col]

    # Use the first Units column found
    if units_columns:
        units_col = units_columns[0]
    else:
        print("WARNING: No Units column found!")
        units_col = None

    # For each row in process
    for idx, row in process.iterrows():
        row_id = row['ID']
        row_career = row['Career']
        py_terms = [row['py1'], row['py2'], row['py3']]

        # For each term column
        for term in sorted_terms:
            if term in py_terms:
                # Term is in py1, py2, or py3 - mark as GS Term
                process.at[idx, term] = "GS Term"
            else:
                # Sum units taken for this ID + Career + Term combination from data
                if units_col:
                    # Build the mask: ID matches AND Career matches
                    mask = (data['ID'] == row_id) & (data['Career'] == row_career)

                    # AND the Term column equals this specific term
                    # The main Term column should be the first one in term_columns
                    if term_columns:
                        main_term_col = term_columns[0]  # Use 'Term' column
                        # Convert term values to string for comparison
                        data_terms = data[main_term_col].fillna('').astype(str)
                        # Remove decimals if present
                        data_terms = data_terms.apply(lambda x: str(int(float(x))) if x and x != '' else '')
                        # Add term match to mask
                        mask = mask & (data_terms == term)

                    # Sum the units taken for matching rows
                    matching_rows = data[mask]
                    total_units = matching_rows[units_col].sum()
                    process.at[idx, term] = total_units
                else:
                    process.at[idx, term] = 0

        # Print progress every 10 rows
        if (idx + 1) % 10 == 0:
            print(f"Processed {idx + 1} rows...")

    # For each row, sum all term columns (skip "GS Term" values, only sum numbers)
    def calculate_lifetime_credits(row):
        total = 0
        for term in sorted_terms:
            value = row[term]
            if value != "GS Term" and value != "":
                try:
                    total += float(value)
                except:
                    pass
        return total

    process['Lifetime Credits'] = process.apply(calculate_lifetime_credits, axis=1)

    # Create placeholder columns for formulas (will be filled when creating workbooks)
    process['Credits to Charge'] = ""
    process['Tuition to Charge'] = ""

    # Define formula creation functions for later use
    def create_credits_formula(row_index):
        excel_row = row_index + 2  # +2 because pandas is 0-indexed and Excel header is row 1
        return f"=IF(W{excel_row}>0, IF(X{excel_row}<7,0, MIN(W{excel_row},X{excel_row}-6)),0)"

    def create_tuition_formula(row_index):
        excel_row = row_index + 2  # +2 because pandas is 0-indexed and Excel header is row 1
        return f'=IF(D{excel_row}="UGRD",Y{excel_row}*{undergraduate_rate},Y{excel_row}*{graduate_rate})'

    # Get current date for filename
    from datetime import datetime
    current_date = datetime.now().strftime("%m.%d.%y")

    # Create Report workbook with unfiltered data first
    print()
    print("Creating Report workbook...")
    report_filename = f"ALL_CAMPUS_{current_date}.xlsx"
    report_path = os.path.join(base_path, report_filename)

    # Create a copy of process for the report and add formulas
    report_process = process.copy().reset_index(drop=True)
    report_process['Credits to Charge'] = [create_credits_formula(i) for i in range(len(report_process))]
    report_process['Tuition to Charge'] = [create_tuition_formula(i) for i in range(len(report_process))]

    # Create new workbook for report
    report_wb = xw.Book()

    # Add Data tab (tab 1)
    report_data_sheet = report_wb.sheets[0]
    report_data_sheet.name = "Data"
    # Write unfiltered data without index
    report_data_sheet.range("A1").options(index=False).value = data
    report_data_sheet.range("A:A").column_width = 105 / 7  # Convert pixels to Excel width units

    # Format Data tab header row
    num_cols_report_data = len(data.columns)
    report_header_range_data = report_data_sheet.range((1, 1), (1, num_cols_report_data))
    report_header_range_data.api.Font.Bold = True
    report_header_range_data.color = (217, 217, 217)  # Grey fill

    # Add Process tab (tab 2)
    report_process_sheet = report_wb.sheets.add("Process", after=report_data_sheet)
    # Write unfiltered process without index
    report_process_sheet.range("A1").options(index=False).value = report_process

    # Format Process tab header row
    num_cols_report_process = len(report_process.columns)
    report_header_range_process = report_process_sheet.range((1, 1), (1, num_cols_report_process))
    report_header_range_process.api.Font.Bold = True
    report_header_range_process.color = (217, 217, 217)  # Grey fill

    # Process tab specific formatting
    report_process_sheet.range("A:A").column_width = 105 / 7  # Convert pixels to Excel width units

    # Hide columns F, G, H
    for col in ['F', 'G', 'H']:
        report_process_sheet.range(f"{col}:{col}").api.EntireColumn.Hidden = True

    # Freeze panes at I2
    report_process_sheet.range("I2").select()
    report_process_sheet.api.Application.ActiveWindow.FreezePanes = True

    # Add SUM formulas at the bottom of columns I:Z
    last_data_row = len(report_process) + 1  # +1 for header
    sum_row = last_data_row + 1

    # Helper function to convert column number to Excel column letter
    def col_num_to_letter(n):
        """Convert 1-indexed column number to Excel column letter (A, B, ..., Z, AA, AB, ...)"""
        result = ""
        while n > 0:
            n -= 1
            result = chr(65 + (n % 26)) + result
            n //= 26
        return result

    # Add SUM formulas for columns I through Z (columns 9-26)
    for col_num in range(9, 27):  # I=9, Z=26
        col_letter = col_num_to_letter(col_num)
        formula = f"=SUM({col_letter}2:{col_letter}{last_data_row})"
        report_process_sheet.range(f"{col_letter}{sum_row}").value = formula

    # Make the sum row bold
    sum_range = report_process_sheet.range(f"I{sum_row}:Z{sum_row}")
    sum_range.api.Font.Bold = True

    # Format column Z as currency with no decimals
    col_z_range = report_process_sheet.range(f"Z2:Z{sum_row}")
    col_z_range.number_format = "$#,##0"

    # Save the Report workbook
    report_wb.save(report_path)
    report_wb.close()

    print(f"  Saved: {report_filename}")
    print()

    # Find all unique campus values
    unique_campuses = process['Campus'].unique()
    print(f"Found {len(unique_campuses)} unique campuses: {list(unique_campuses)}")
    print()

    # Create a workbook for each campus
    for campus in unique_campuses:

        # Filter data and process by campus
        campus_data = data[data['Campus'] == campus].copy()
        campus_process = process[process['Campus'] == campus].copy().reset_index(drop=True)

        # Add formulas with correct row numbers for this campus
        campus_process['Credits to Charge'] = [create_credits_formula(i) for i in range(len(campus_process))]
        campus_process['Tuition to Charge'] = [create_tuition_formula(i) for i in range(len(campus_process))]

        # Create filename with campus name and date
        output_filename = f"{campus}_{current_date}.xlsx"
        output_path = os.path.join(base_path, output_filename)

        # Create new workbook
        new_wb = xw.Book()

        # Add Data tab (tab 1)
        data_sheet = new_wb.sheets[0]
        data_sheet.name = "Data"
        # Write data without index
        data_sheet.range("A1").options(index=False).value = campus_data
        data_sheet.range("A:A").column_width = 105 / 7  # Convert pixels to Excel width units (approx)


        # Format Data tab header row
        # Get the number of columns from the dataframe instead of using end('right')
        num_cols_data = len(campus_data.columns)
        header_range_data = data_sheet.range((1, 1), (1, num_cols_data))
        header_range_data.api.Font.Bold = True
        header_range_data.color = (217, 217, 217)  # Grey fill

        # Add Process tab (tab 2)
        process_sheet = new_wb.sheets.add("Process", after=data_sheet)
        # Write data without index
        process_sheet.range("A1").options(index=False).value = campus_process

        # Format Process tab header row
        # Get the number of columns from the dataframe instead of using end('right')
        num_cols_process = len(campus_process.columns)
        header_range_process = process_sheet.range((1, 1), (1, num_cols_process))
        header_range_process.api.Font.Bold = True
        header_range_process.color = (217, 217, 217)  # Grey fill

        # Process tab specific formatting
        process_sheet.range("A:A").column_width = 105 / 7  # Convert pixels to Excel width units (approx)

        # Hide columns
        for col in ['F', 'G', 'H']:
            process_sheet.range(f"{col}:{col}").api.EntireColumn.Hidden = True

        # Freeze panes
        process_sheet.range("I2").select()
        process_sheet.api.Application.ActiveWindow.FreezePanes = True

        # Add SUM formulas at the bottom of columns I:Z
        campus_last_data_row = len(campus_process) + 1  # +1 for header
        campus_sum_row = campus_last_data_row + 1

        # Add SUM formulas for columns I through Z (columns 9-26)
        for col_num in range(9, 27):  # I=9, Z=26
            col_letter = col_num_to_letter(col_num)
            formula = f"=SUM({col_letter}2:{col_letter}{campus_last_data_row})"
            process_sheet.range(f"{col_letter}{campus_sum_row}").value = formula

        # Make the sum row bold
        campus_sum_range = process_sheet.range(f"I{campus_sum_row}:Z{campus_sum_row}")
        campus_sum_range.api.Font.Bold = True

        # Format column Z as currency with no decimals
        campus_col_z_range = process_sheet.range(f"Z2:Z{campus_sum_row}")
        campus_col_z_range.number_format = "$#,##0"

        # Save the workbook
        new_wb.save(output_path)
        new_wb.close()

        print(f"  Saved: {output_filename}")
        print()

    # Close original workbook
    wb.close()
    app.quit()

    print("All campus workbooks created successfully!")
    print()

except Exception as e:
    print("\n" + "="*80)
    print("ERROR OCCURRED:")
    print("="*80)
    print(f"\nError type: {type(e).__name__}")
    print(f"Error message: {str(e)}")
    print("\nFull traceback:")
    traceback.print_exc()
    print("="*80)

    # Try to close workbooks if they're open
    try:
        if 'wb' in locals():
            wb.close()
        if 'app' in locals():
            app.quit()
    except:
        pass

finally:
    # Wait for user input before closing
    input("\nPress Enter to exit...")


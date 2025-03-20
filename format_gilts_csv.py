#!/usr/bin/env python3
"""
Script to convert the downloaded Excel file to a properly formatted CSV.
This script handles specific formatting requirements for the Gilts data.
"""
import os
import sys
import csv
import xlrd
from datetime import datetime
import glob
import re

def format_gilts_csv(excel_file=None, output_dir=None):
    """
    Convert Excel file to CSV with proper formatting:
    - Row 1 and 6 of XLS become rows 1 and 2 of CSV
    - Row 9 of XLS becomes row 3 of CSV (headers)
    - Handle dynamic section headers and notes
    
    This function dynamically identifies the structure of the Excel file and formats it according
    to the specified requirements. It handles both conventional gilts and index-linked gilts sections.
    
    Args:
        excel_file: Path to the Excel file to convert. If None, will use the latest file in downloads directory.
        output_dir: Directory to save the CSV file. Defaults to 'csv_exports' directory.
    
    Returns:
        Path to the created CSV file or None if conversion failed.
    """
    # Set default output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'csv_exports')
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # If no file is specified, find the most recent Excel file in the downloads directory
    if excel_file is None:
        downloads_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'downloads')
        excel_files = glob.glob(os.path.join(downloads_dir, 'gilts_in_issue_*.xls*'))
        
        if not excel_files:
            print("Error: No Excel files found in the downloads directory.")
            return None
        
        # Get the most recent file
        excel_file = max(excel_files, key=os.path.getmtime)
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Error: File not found at {excel_file}")
        return None
    
    print(f"Converting file: {excel_file}")
    
    try:
        # Extract date from filename or use current date
        try:
            # Try to extract date from filename (format: gilts_in_issue_DD-MM-YYYY.xls)
            filename = os.path.basename(excel_file)
            date_str = filename.split('_')[-1].split('.')[0]  # Extract DD-MM-YYYY
            date_obj = datetime.strptime(date_str, '%d-%m-%Y')
            formatted_date = date_obj.strftime('%Y%m%d')
        except (ValueError, IndexError):
            # If date extraction fails, use current date
            formatted_date = datetime.now().strftime('%Y%m%d')
        
        # Create CSV filename with date
        csv_filename = f"gilts_in_issue_{formatted_date}.csv"
        csv_path = os.path.join(output_dir, csv_filename)
        
        # Open the workbook
        workbook = xlrd.open_workbook(excel_file)
        sheet = workbook.sheet_by_index(0)
        
        print(f"Excel file has {workbook.nsheets} sheets")
        print(f"Sheet dimensions: {sheet.nrows} rows, {sheet.ncols} columns")
        
        # Find the row with headers (should be row 9 in the Excel file, index 8)
        header_row_idx = None
        for i in range(min(15, sheet.nrows)):  # Look in first 15 rows
            row_values = [str(sheet.cell_value(i, j)).strip() for j in range(sheet.ncols)]
            # Check if this row contains typical headers
            if 'ISIN Code' in row_values and 'Redemption Date' in row_values:
                header_row_idx = i
                print(f"Found header row at index {i} (row {i+1} in Excel)")
                break
        
        if header_row_idx is None:
            print("Error: Could not find header row in Excel file")
            return None
        
        # Identify section headers
        section_headers = []
        for i in range(header_row_idx + 1, sheet.nrows):
            cell_value = str(sheet.cell_value(i, 0)).strip()
            
            # Skip group labels (Ultra-Short, Short, Medium, Long)
            if cell_value in ['Ultra-Short', 'Short', 'Medium', 'Long']:
                continue
            
            # Check if this might be a section header
            # Look for rows with content in first column but no ISIN code in second column
            # Or rows that contain "Index-linked" which indicate a new section
            if (cell_value and not sheet.cell_value(i, 1) and i > header_row_idx + 1) or \
               ('Index-linked' in cell_value and i > header_row_idx + 1):
                # This is likely a section header
                section_headers.append(i)
        
        print(f"Found section headers at rows: {[idx+1 for idx in section_headers]}")
        
        # Write to CSV
        with open(csv_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            
            # Row 1: Title row from Excel
            writer.writerow([str(sheet.cell_value(0, j)).strip() for j in range(sheet.ncols)])
            
            # Row 2: Total amount row from Excel (row 6, index 5)
            writer.writerow([str(sheet.cell_value(5, j)).strip() for j in range(sheet.ncols)])
            
            # Row 3: Headers from Excel
            headers = [str(sheet.cell_value(header_row_idx, j)).strip() for j in range(sheet.ncols)]
            writer.writerow(headers)
            
            # Process data rows
            for i in range(header_row_idx + 1, sheet.nrows):
                row_values = [str(sheet.cell_value(i, j)).strip() for j in range(sheet.ncols)]
                
                # Skip empty rows
                if not any(row_values):
                    continue
                
                # Skip group label rows (Ultra-Short, Short, Medium, Long)
                if row_values[0] in ['Ultra-Short', 'Short', 'Medium', 'Long']:
                    continue
                
                # Write all other rows as-is
                writer.writerow(row_values)
        
        print(f"Successfully converted to CSV: {csv_path}")
        return csv_path
    
    except Exception as e:
        print(f"Error converting file: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Get Excel file path from command line argument or use latest file
    excel_file = sys.argv[1] if len(sys.argv) > 1 else None
    
    # Convert to CSV
    csv_path = format_gilts_csv(excel_file)
    
    if csv_path:
        print(f"\nCSV file created at: {csv_path}")
    else:
        print("\nFailed to create CSV file.")

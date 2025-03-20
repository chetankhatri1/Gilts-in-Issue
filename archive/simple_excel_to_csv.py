#!/usr/bin/env python3
"""
Simple script to convert the downloaded Excel file to a CSV with date in the filename.
This script uses xlrd directly for older Excel formats.
"""
import os
import sys
import csv
import xlrd
from datetime import datetime
import glob

def excel_to_csv(excel_file=None, output_dir=None):
    """
    Convert Excel file to CSV with date in the filename using xlrd directly.
    
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
            # Try to extract date from filename (format: gilts_in_issue_DD-MM-YYYY.xlsx)
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
        
        # Find the actual data rows (skip header rows)
        start_row = 0
        for i in range(min(30, sheet.nrows)):
            row_values = [str(sheet.cell_value(i, j)).strip() for j in range(sheet.ncols)]
            print(f"Row {i+1}: {row_values}")
            # Look for rows that might contain headers like ISIN, Name, etc.
            if any(keyword in ' '.join(row_values).upper() for keyword in ['ISIN', 'GILT', 'COUPON', 'MATURITY']):
                start_row = i
                print(f"Found potential header row at row {i+1}")
                break
        
        # Write to CSV
        with open(csv_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write all rows from the Excel file
            for i in range(sheet.nrows):
                row = [sheet.cell_value(i, j) for j in range(sheet.ncols)]
                # Skip completely empty rows
                if any(str(cell).strip() for cell in row):
                    writer.writerow(row)
        
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
    csv_path = excel_to_csv(excel_file)
    
    if csv_path:
        print(f"\nCSV file created at: {csv_path}")
    else:
        print("\nFailed to create CSV file.")

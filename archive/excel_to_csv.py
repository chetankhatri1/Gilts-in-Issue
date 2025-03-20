#!/usr/bin/env python3
"""
Script to convert the downloaded Excel file to a CSV with date in the filename.
"""
import os
import sys
import pandas as pd
from datetime import datetime
import glob

def excel_to_csv(excel_file=None, output_dir=None):
    """
    Convert Excel file to CSV with date in the filename.
    
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
        excel_files = glob.glob(os.path.join(downloads_dir, 'gilts_in_issue_*.xlsx'))
        
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
        # First try to get basic info about the file
        import xlrd
        workbook = xlrd.open_workbook(excel_file)
        print(f"Excel file has {workbook.nsheets} sheets")
        
        # For older .xls files, we'll use xlrd directly for more control
        if excel_file.lower().endswith('.xls'):
            # Get the first sheet
            sheet = workbook.sheet_by_index(0)
            
            # Find the data rows (skip header rows)
            start_row = 0
            for i in range(min(20, sheet.nrows)):
                row_values = [str(sheet.cell_value(i, j)).strip() for j in range(sheet.ncols)]
                if any('ISIN' in val for val in row_values):
                    start_row = i
                    break
            
            # Create a DataFrame from the data
            data = []
            headers = []
            
            # Get headers
            if start_row < sheet.nrows:
                headers = [str(sheet.cell_value(start_row, j)).strip() for j in range(sheet.ncols)]
                headers = [h if h else f"Column_{j}" for j, h in enumerate(headers)]
            
            # Get data rows
            for i in range(start_row + 1, sheet.nrows):
                row = [sheet.cell_value(i, j) for j in range(sheet.ncols)]
                if any(row):  # Skip empty rows
                    data.append(row)
            
            # Create DataFrame
            df = pd.DataFrame(data, columns=headers)
        else:
            # For newer .xlsx files
            df = pd.read_excel(excel_file, engine='openpyxl')
            
        # Clean the DataFrame
        # Remove empty columns
        df = df.dropna(axis=1, how='all')
        # Remove empty rows
        df = df.dropna(axis=0, how='all')
        
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
        
        # Save to CSV
        df.to_csv(csv_path, index=False)
        print(f"Successfully converted to CSV: {csv_path}")
        
        # Display basic info about the data
        print(f"\nData shape: {df.shape} (rows, columns)")
        print("\nFirst few rows:")
        print(df.head())
        
        return csv_path
    
    except Exception as e:
        print(f"Error converting file: {e}")
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

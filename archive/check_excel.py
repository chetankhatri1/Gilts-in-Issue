#!/usr/bin/env python3
"""
Simple script to check Excel file content using multiple libraries.
"""
import os
import sys
import pandas as pd

def check_file(file_path):
    """Check file using multiple methods."""
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return
    
    print(f"Checking file: {file_path}")
    file_size = os.path.getsize(file_path)
    print(f"File size: {file_size} bytes")
    
    # Try different methods to read the file
    try:
        print("\nAttempting to read with xlrd...")
        import xlrd
        workbook = xlrd.open_workbook(file_path)
        print(f"Success! Found {workbook.nsheets} sheets:")
        for sheet_idx in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_idx)
            print(f"  - Sheet {sheet_idx+1}: {sheet.name} ({sheet.nrows} rows, {sheet.ncols} columns)")
            
            # Print first few rows of first sheet
            if sheet_idx == 0 and sheet.nrows > 0:
                print("\nFirst few rows of first sheet:")
                for row_idx in range(min(5, sheet.nrows)):
                    print(f"Row {row_idx+1}: {[sheet.cell_value(row_idx, col_idx) for col_idx in range(sheet.ncols)]}")
    except Exception as e:
        print(f"Error with xlrd: {e}")
    
    try:
        print("\nAttempting to read with pandas/xlrd...")
        df = pd.read_excel(file_path, engine='xlrd')
        print("Success! DataFrame shape:", df.shape)
        print("\nFirst few rows:")
        print(df.head())
        print("\nColumns:", df.columns.tolist())
    except Exception as e:
        print(f"Error with pandas/xlrd: {e}")
    
    try:
        print("\nAttempting to read with pandas/openpyxl...")
        df = pd.read_excel(file_path, engine='openpyxl')
        print("Success! DataFrame shape:", df.shape)
        print("\nFirst few rows:")
        print(df.head())
        print("\nColumns:", df.columns.tolist())
    except Exception as e:
        print(f"Error with pandas/openpyxl: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 check_excel.py <path_to_excel_file>")
        sys.exit(1)
    
    check_file(sys.argv[1])

#!/usr/bin/env python3
"""
Process Gilts in Issue data from manually downloaded Excel files.
This script helps organize and analyze UK Gilts data after manual download.
"""
import os
import sys
import pandas as pd
from datetime import datetime
import shutil

def process_gilts_file(file_path):
    """
    Process a manually downloaded Gilts in Issue Excel file.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        pandas.DataFrame: Processed data
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return None
    
    print(f"Processing file: {file_path}")
    
    try:
        # Try to read the Excel file - determine engine based on file extension
        if file_path.endswith('.xls'):
            # For older .xls files
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            # For newer .xlsx files
            df = pd.read_excel(file_path, engine='openpyxl')
        
        # Check if it's a valid Gilts file by looking for expected columns
        expected_columns = ['ISIN', 'Name', 'Coupon', 'Maturity Date']
        found_columns = [col for col in expected_columns if any(col.lower() in str(c).lower() for c in df.columns)]
        
        if not found_columns:
            print("Warning: This doesn't appear to be a valid Gilts in Issue file.")
            print(f"Found columns: {df.columns.tolist()}")
            return df
        
        # Basic data cleaning
        # Remove any completely empty rows or columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Display basic information
        print("\nFile successfully processed!")
        print(f"Shape: {df.shape}")
        print(f"Columns: {df.columns.tolist()}")
        print("\nFirst few rows:")
        print(df.head())
        
        return df
    
    except Exception as e:
        print(f"Error processing file: {e}")
        return None

def organize_files(source_path, date_str=None):
    """
    Organize downloaded files by moving them to the downloads directory
    with proper naming convention.
    
    Args:
        source_path (str): Path to the source file
        date_str (str, optional): Date string in DD/MM/YYYY format
        
    Returns:
        str: Path to the organized file
    """
    # Create downloads directory if it doesn't exist
    script_dir = os.path.dirname(os.path.abspath(__file__))
    downloads_dir = os.path.join(script_dir, "downloads")
    os.makedirs(downloads_dir, exist_ok=True)
    
    # Generate a filename with date
    if date_str:
        try:
            date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            file_date = date_obj.strftime('%d-%m-%Y')
        except ValueError:
            print("Invalid date format. Using current date.")
            file_date = datetime.now().strftime('%d-%m-%Y')
    else:
        file_date = datetime.now().strftime('%d-%m-%Y')
    
    # Get file extension
    _, ext = os.path.splitext(source_path)
    
    # Create destination path
    dest_path = os.path.join(downloads_dir, f"gilts_in_issue_{file_date}{ext}")
    
    # Copy the file
    try:
        shutil.copy2(source_path, dest_path)
        print(f"File copied to: {dest_path}")
        return dest_path
    except Exception as e:
        print(f"Error copying file: {e}")
        return source_path

def main():
    """Main function to process Gilts data files."""
    print("UK Gilts in Issue Data Processor")
    print("================================")
    
    if len(sys.argv) < 2:
        print("\nUsage: python3 process_gilts_data.py <path_to_excel_file> [date_in_DD/MM/YYYY]")
        print("\nExample: python3 process_gilts_data.py ~/Downloads/D1A.xlsx 19/03/2025")
        print("\nIf date is not provided, current date will be used for organizing files.")
        
        # Ask for file path interactively
        file_path = input("\nEnter the path to the downloaded Excel file: ")
        date_str = input("Enter the date of the data (DD/MM/YYYY) or press Enter for today's date: ")
        
        if not file_path:
            print("No file path provided. Exiting.")
            return
    else:
        file_path = sys.argv[1]
        date_str = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Organize the file
    organized_path = organize_files(file_path, date_str)
    
    # Process the file
    df = process_gilts_file(organized_path)
    
    if df is not None:
        # Ask if user wants to save as CSV
        save_csv = input("\nDo you want to save the data as CSV? (y/n): ").lower()
        if save_csv == 'y':
            csv_path = organized_path.replace('.xlsx', '.csv')
            df.to_csv(csv_path, index=False)
            print(f"Data saved to CSV: {csv_path}")

if __name__ == "__main__":
    main()

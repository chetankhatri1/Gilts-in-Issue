#!/usr/bin/env python3
"""
Combined script to download and format Gilts in Issue data.
This script handles the entire process from downloading the Excel file from the DMO website
to converting it to a properly formatted CSV file.
"""
import os
import sys
import csv
import time
import random
import xlrd
import asyncio
import traceback
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

async def download_gilts_data(date_str=None):
    """
    Download Gilts in Issue data from the UK Debt Management Office website using Playwright.
    
    Args:
        date_str (str, optional): Date in format 'DD/MM/YYYY'. If None, yesterday's date is used.
    
    Returns:
        str: Path to the downloaded file
    """
    # Base URL for the Gilts in Issue data
    base_url = "https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D1A"
    
    # If no date provided, use yesterday's date
    if not date_str:
        yesterday = datetime.now() - timedelta(days=1)
        date_str = yesterday.strftime('%d/%m/%Y')
    
    print(f"Attempting to download Gilts data for date: {date_str}")
    
    # Create output directory if it doesn't exist
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "downloads")
    os.makedirs(output_dir, exist_ok=True)
    
    async with async_playwright() as p:
        # Launch a browser with a realistic viewport and user agent
        browser_type = p.chromium  # Can also use p.firefox or p.webkit
        
        # Use a browser context with various options to appear more human-like
        browser = await browser_type.launch(headless=False)  # Set to True for headless mode
        
        # Create a context with specific options
        context = await browser.new_context(
            viewport={'width': 1280, 'height': 800},
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
            accept_downloads=True,
            locale='en-GB',
            timezone_id='Europe/London',
            # Add additional human-like characteristics
            device_scale_factor=1.0,
            has_touch=False,
            is_mobile=False
        )
        
        # Add some human-like behaviors
        # 1. Enable JavaScript and cookies
        await context.add_cookies([{
            'name': 'user_preference',
            'value': 'accepted',
            'domain': 'www.dmo.gov.uk',
            'path': '/'
        }])
        
        # Create a new page
        page = await context.new_page()
        
        try:
            # Navigate to the DMO website with a timeout
            print("Navigating to DMO website...")
            await page.goto(base_url, wait_until='networkidle', timeout=60000)
            
            # Add a random delay to mimic human behavior
            await asyncio.sleep(random.uniform(1.5, 3.0))
            
            # Handle cookie popup if it appears
            print("Checking for cookie popup...")
            try:
                # Look for common cookie consent buttons
                cookie_selectors = [
                    "button:has-text('Accept')", 
                    "button:has-text('Accept All')",
                    "button:has-text('Accept Cookies')",
                    "button:has-text('I Accept')",
                    "button:has-text('OK')",
                    "button:has-text('Agree')",
                    "#onetrust-accept-btn-handler",
                    ".cookie-accept-button",
                    ".accept-cookies"
                ]
                
                for selector in cookie_selectors:
                    cookie_button = await page.query_selector(selector)
                    if cookie_button:
                        print(f"Found cookie popup, clicking {selector}...")
                        await cookie_button.click()
                        print("Accepted cookies")
                        # Wait for popup to disappear
                        await asyncio.sleep(random.uniform(1.0, 2.0))
                        break
            except Exception as e:
                print(f"Error handling cookie popup: {e}")
            
            # Debug: Let's check what form elements are available
            print("Analyzing page elements...")
            form_elements = await page.evaluate("""
                () => {
                    const inputs = Array.from(document.querySelectorAll('input'));
                    return inputs.map(input => {
                        return {
                            id: input.id,
                            name: input.name,
                            type: input.type,
                            placeholder: input.placeholder
                        };
                    });
                }
            """)
            print(f"Found form elements: {form_elements}")
            
            # Skip setting the date as it's correct by default
            print("Using default date...")
            
            # Add a random delay to mimic human behavior
            await asyncio.sleep(random.uniform(1.0, 2.0))
            
            # Set up download handler
            download_path = None
            
            async def handle_download(download):
                nonlocal download_path
                print("Download started...")
                # Wait for the download to complete
                download_path = await download.path()
                print(f"Download completed: {download_path}")
            
            # Listen for download events
            page.on('download', handle_download)
            
            # Debug: Let's find all buttons on the page
            print("Finding Excel button...")
            buttons = await page.evaluate("""
                () => {
                    const buttons = Array.from(document.querySelectorAll('button, input[type="button"], a.button, .btn'));
                    return buttons.map(button => {
                        return {
                            text: button.innerText || button.value,
                            id: button.id,
                            class: button.className
                        };
                    });
                }
            """)
            print(f"Found buttons: {buttons}")
                
            # Find and click the Excel button
            print("Clicking Excel button...")
            excel_button = await page.query_selector('button:has-text("Excel")')
            
            if not excel_button:
                print("Trying alternative Excel button selectors...")
                excel_selectors = [
                    'input[value*="Excel"]',
                    'a:has-text("Excel")',
                    '.btn:has-text("Excel")',
                    'button[id*="excel"]',
                    'button[class*="excel"]'
                ]
                
                for selector in excel_selectors:
                    print(f"Trying Excel button selector: {selector}")
                    excel_button = await page.query_selector(selector)
                    if excel_button:
                        print(f"Found Excel button with selector: {selector}")
                        break
                
            if excel_button:
                # Move mouse to button with human-like motion
                await page.mouse.move(
                    random.uniform(0, 100), 
                    random.uniform(0, 100), 
                    steps=random.randint(5, 10)
                )
                
                # Click the button
                await excel_button.click()
                print("Clicked Excel button")
                
                # Wait for download to start and complete
                print("Waiting for download to complete...")
                # Wait up to 30 seconds for the download to complete
                download_wait_time = 0
                while download_path is None and download_wait_time < 30:
                    await asyncio.sleep(1)
                    download_wait_time += 1
                
                if download_path:
                    # Copy the file to our downloads directory with proper naming
                    file_date = date_str.replace('/', '-')
                    output_file = os.path.join(output_dir, f"gilts_in_issue_{file_date}.xls")
                    
                    # If the file exists, remove it first
                    if os.path.exists(output_file):
                        os.remove(output_file)
                    
                    # Copy the downloaded file
                    os.rename(download_path, output_file)
                    print(f"Successfully downloaded Gilts data to: {output_file}")
                    return output_file
                else:
                    print("Download failed or timed out")
            else:
                print("Could not find Excel button")
        
        except Exception as e:
            print(f"Error during download: {e}")
        finally:
            # Close the browser
            await browser.close()
    
    return None

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
        
        print(f"Found section headers at rows: {section_headers}")
        
        # Write to CSV
        with open(csv_path, 'w', newline='') as csvfile:
            # Use csv.QUOTE_ALL to ensure all fields are quoted
            writer = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
            
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
                
                # Write all other rows with quoting
                writer.writerow(row_values)
        
        print(f"Successfully converted to CSV: {csv_path}")
        return csv_path
    
    except Exception as e:
        print(f"Error converting file: {e}")
        traceback.print_exc()
        return None

async def main():
    """Main function to download and format Gilts data."""
    # Calculate yesterday's date
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_str = yesterday.strftime('%d/%m/%Y')
    
    try:
        # Step 1: Download the Excel file
        print("\n=== STEP 1: DOWNLOADING GILTS DATA ===\n")
        excel_file = await download_gilts_data(yesterday_str)
        
        if excel_file and os.path.exists(excel_file):
            # Try to verify if we got a real Excel file
            file_type = os.popen(f"file '{excel_file}'").read().strip()
            print(f"File type: {file_type}")
            
            if "HTML" in file_type or "html" in file_type:
                print("Warning: Downloaded file appears to be HTML, not Excel.")
                print("Bot protection may still be active.")
                return
            elif "Excel" in file_type or "Microsoft" in file_type or "Zip archive" in file_type or "Composite Document File" in file_type:
                print("Success! Downloaded a valid Excel file.")
            
            # Step 2: Format the Excel file to CSV
            print("\n=== STEP 2: FORMATTING TO CSV ===\n")
            csv_path = format_gilts_csv(excel_file)
            
            if csv_path:
                print(f"\nComplete process successful!")
                print(f"Downloaded Excel file: {excel_file}")
                print(f"Formatted CSV file: {csv_path}")
            else:
                print("\nCSV formatting failed.")
        else:
            print("\nDownload failed. Please try manual download:")
            print(f"1. Visit: https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D1A")
            print(f"2. Enter date: {yesterday_str}")
            print("3. Click 'Excel' button")
    except Exception as e:
        print(f"Error in main process: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    # Add missing import for glob
    import glob
    
    # Run the main async function
    asyncio.run(main())

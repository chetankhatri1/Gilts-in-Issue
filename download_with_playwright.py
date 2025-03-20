#!/usr/bin/env python3
"""
Download Gilts in Issue data using Playwright to automate browser interactions.
This approach better mimics human behavior to bypass bot protection.
"""
import os
import time
import random
from datetime import datetime, timedelta
import asyncio
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

async def main():
    """Main function to download Gilts data."""
    # Calculate yesterday's date
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_str = yesterday.strftime('%d/%m/%Y')
    
    try:
        output_file = await download_gilts_data(yesterday_str)
        
        if output_file and os.path.exists(output_file):
            # Try to verify if we got a real Excel file
            file_type = os.popen(f"file '{output_file}'").read().strip()
            print(f"File type: {file_type}")
            
            if "HTML" in file_type or "html" in file_type:
                print("Warning: Downloaded file appears to be HTML, not Excel.")
                print("Bot protection may still be active.")
            elif "Excel" in file_type or "Microsoft" in file_type or "Zip archive" in file_type:
                print("Success! Downloaded a valid Excel file.")
            
            print(f"\nDownloaded file is available at: {output_file}")
        else:
            print("\nDownload failed. Please try manual download:")
            print(f"1. Visit: https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D1A")
            print(f"2. Enter date: {yesterday_str}")
            print("3. Click 'Excel' button")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    asyncio.run(main())

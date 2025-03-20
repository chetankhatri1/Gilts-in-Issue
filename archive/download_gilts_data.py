#!/usr/bin/env python3
import os
import time
import random
import requests
from datetime import datetime, timedelta
import pandas as pd

def download_gilts_data(date_str=None):
    """
    Download Gilts in Issue data from the UK Debt Management Office website using requests.
    
    Args:
        date_str (str, optional): Date in format 'DD/MM/YYYY'. If None, yesterday's date is used.
    
    Returns:
        str: Path to the downloaded Excel file
    """
    # Base URL for the Gilts in Issue data
    base_url = "https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D1A"
    
    # If no date provided, use yesterday's date
    if not date_str:
        yesterday = datetime.now() - timedelta(days=1)
        date_str = yesterday.strftime('%d/%m/%Y')
    
    print(f"Attempting to download Gilts data for date: {date_str}")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "downloads")
    os.makedirs(output_dir, exist_ok=True)
    
    # Common browser user agents
    user_agents = [
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0"
    ]
    
    # Select a random user agent
    user_agent = random.choice(user_agents)
    
    # Set headers to mimic a browser
    headers = {
        "User-Agent": user_agent,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Sec-Fetch-User": "?1",
        "Cache-Control": "max-age=0"
    }
    
    # Create a session to maintain cookies
    session = requests.Session()
    
    try:
        # First, get the page to obtain any necessary cookies or tokens
        print("Accessing DMO website...")
        response = session.get(base_url, headers=headers)
        response.raise_for_status()  # Raise an exception for HTTP errors
        
        # Add a delay to mimic human behavior
        time.sleep(random.uniform(1, 3))
        
        # Prepare the form data for Excel download
        form_data = {
            "reportCode": "D1A",
            "date": date_str,
            "format": "Excel"
        }
        
        # Add a referer header to make the request look more legitimate
        headers["Referer"] = base_url
        
        # Submit the form to download the Excel file
        print(f"Requesting Excel file for date: {date_str}...")
        download_url = "https://www.dmo.gov.uk/data/ExcelDataReport"
        response = session.post(download_url, data=form_data, headers=headers, allow_redirects=True)
        response.raise_for_status()
        
        # Save the Excel file
        file_date = date_str.replace('/', '-')
        output_file = os.path.join(output_dir, f"gilts_in_issue_{file_date}.xlsx")
        
        with open(output_file, 'wb') as f:
            f.write(response.content)
        
        print(f"Successfully downloaded Gilts data to: {output_file}")
        return output_file
        
    except requests.exceptions.RequestException as e:
        print(f"Error during download: {e}")
        
        # Try alternative direct URL approach
        try:
            print("Attempting alternative download method...")
            # Format date for URL
            url_date = datetime.strptime(date_str, '%d/%m/%Y').strftime('%Y-%m-%d')
            direct_url = f"https://www.dmo.gov.uk/data/ExcelDataReport?reportCode=D1A&date={url_date}"
            
            response = session.get(direct_url, headers=headers)
            response.raise_for_status()
            
            output_file = os.path.join(output_dir, f"gilts_in_issue_{file_date}.xlsx")
            with open(output_file, 'wb') as f:
                f.write(response.content)
            
            print(f"Alternative download successful: {output_file}")
            return output_file
        except requests.exceptions.RequestException as alt_e:
            print(f"Alternative download method also failed: {alt_e}")
            raise

if __name__ == "__main__":
    # Calculate yesterday's date
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_str = yesterday.strftime('%d/%m/%Y')
    
    try:
        output_file = download_gilts_data(yesterday_str)
        
        # Try to read and display some basic info from the Excel file
        try:
            df = pd.read_excel(output_file, engine='openpyxl')
            print(f"\nSuccessfully loaded Excel file. Preview of data:")
            print(f"Shape: {df.shape}")
            print(f"Columns: {df.columns.tolist()}")
            print("\nFirst few rows:")
            print(df.head())
        except Exception as e:
            print(f"Note: Could not parse Excel file for preview: {e}")
            print(f"You can still open the file manually at: {output_file}")
            
    except Exception as e:
        print(f"Error downloading Gilts data: {e}")
        print("\nTroubleshooting tips:")
        print("1. Check your internet connection")
        print("2. The DMO website might be temporarily unavailable")
        print("3. The website might have changed its structure")
        print("4. Try visiting the URL manually: https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D1A")

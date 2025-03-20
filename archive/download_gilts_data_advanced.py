#!/usr/bin/env python3
import os
import time
import random
import json
from datetime import datetime, timedelta
from requests_html import HTMLSession
import pandas as pd

def download_gilts_data(date_str=None):
    """
    Download Gilts in Issue data from the UK Debt Management Office website using requests-html.
    
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
    
    # Create a session
    session = HTMLSession()
    session.headers.update({
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
    })
    
    try:
        # First, get the page and render it (this executes JavaScript)
        print("Accessing DMO website and rendering JavaScript...")
        response = session.get(base_url)
        response.html.render(sleep=3, timeout=20)  # Render JavaScript and wait
        
        # Add a delay to mimic human behavior
        time.sleep(random.uniform(2, 4))
        
        # Find the date input field and set the date
        print(f"Setting date to: {date_str}")
        date_script = f"""
        document.getElementById('date').value = '{date_str}';
        """
        response.html.render(script=date_script, reload=False)
        
        # Add another delay to mimic human behavior
        time.sleep(random.uniform(1, 2))
        
        # Now prepare to click the Excel button
        # First, we'll get the form data
        form_data = {
            "reportCode": "D1A",
            "date": date_str,
            "format": "Excel"
        }
        
        # Add a referer header
        session.headers.update({"Referer": base_url})
        
        # Submit the form to download the Excel file
        print("Requesting Excel file...")
        download_url = "https://www.dmo.gov.uk/data/ExcelDataReport"
        excel_response = session.post(download_url, data=form_data, allow_redirects=True)
        
        # Check if we got an Excel file or HTML
        content_type = excel_response.headers.get('Content-Type', '')
        if 'html' in content_type.lower():
            print("Warning: Received HTML instead of Excel. Bot protection may still be active.")
            
            # Try to save the HTML for inspection
            html_file = os.path.join(output_dir, "response.html")
            with open(html_file, 'wb') as f:
                f.write(excel_response.content)
            print(f"Saved HTML response to: {html_file}")
            
            # Try to extract any useful information from the HTML
            try:
                excel_response.html.render(sleep=2)
                print("HTML Content Preview:")
                print(excel_response.html.text[:500])
            except Exception as render_e:
                print(f"Could not render HTML: {render_e}")
                
            # Try an alternative approach - direct API call
            print("\nAttempting direct API call...")
            api_url = "https://www.dmo.gov.uk/data/api/gilts/conventional"
            api_headers = session.headers.copy()
            api_headers.update({
                "Accept": "application/json",
                "Content-Type": "application/json"
            })
            
            try:
                api_response = session.get(api_url, headers=api_headers)
                if api_response.status_code == 200:
                    # Try to parse as JSON
                    try:
                        data = api_response.json()
                        # Save as JSON file
                        json_file = os.path.join(output_dir, f"gilts_data_{date_str.replace('/', '-')}.json")
                        with open(json_file, 'w') as f:
                            json.dump(data, f, indent=2)
                        print(f"Successfully downloaded Gilts data as JSON: {json_file}")
                        
                        # Try to convert to Excel
                        try:
                            df = pd.DataFrame(data)
                            excel_file = os.path.join(output_dir, f"gilts_in_issue_{date_str.replace('/', '-')}.xlsx")
                            df.to_excel(excel_file, index=False)
                            print(f"Converted JSON to Excel: {excel_file}")
                            return excel_file
                        except Exception as excel_e:
                            print(f"Could not convert JSON to Excel: {excel_e}")
                            return json_file
                    except Exception as json_e:
                        print(f"Could not parse API response as JSON: {json_e}")
                else:
                    print(f"API request failed with status code: {api_response.status_code}")
            except Exception as api_e:
                print(f"API request failed: {api_e}")
                
            # If all else fails, suggest manual download
            print("\nAutomated download failed. Please try manual download:")
            print(f"1. Visit: {base_url}")
            print(f"2. Enter date: {date_str}")
            print("3. Click 'Excel' button")
            
            raise Exception("Could not bypass bot protection")
        else:
            # We got a non-HTML response, hopefully an Excel file
            file_date = date_str.replace('/', '-')
            output_file = os.path.join(output_dir, f"gilts_in_issue_{file_date}.xlsx")
            
            with open(output_file, 'wb') as f:
                f.write(excel_response.content)
            
            print(f"Successfully downloaded Gilts data to: {output_file}")
            return output_file
            
    except Exception as e:
        print(f"Error during download: {e}")
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
        print("\nAlternative approach:")
        print("1. Open a web browser and visit: https://www.dmo.gov.uk/data/pdfdatareport?reportCode=D1A")
        print(f"2. Enter yesterday's date ({yesterday_str}) in the date field")
        print("3. Click the 'Excel' button to download the file")

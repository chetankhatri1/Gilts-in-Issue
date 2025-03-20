# UK Gilts in Issue Data Tools

This repository contains tools for downloading and processing Gilts in Issue data from the UK Debt Management Office (DMO) website.

## Features

- Successfully downloads Gilts in Issue data in Excel format using Playwright
- Handles cookie consent popups automatically
- Provides tools for processing downloaded data
- Organizes files with proper naming conventions
- Creates a downloads folder to store the Excel files
- Provides basic information about the downloaded data

## Requirements

Install the required dependencies:

```bash
pip install -r requirements.txt
python -m playwright install
```

## Usage

### Recommended: Automated Download with Playwright

Use Playwright for reliable automated download:

```bash
python3 download_with_playwright.py
```

This script successfully:

1. Navigates to the DMO website
2. Accepts the cookie popup
3. Clicks the Excel button to download the data
4. Saves the file to the `downloads` directory

### Alternative Approaches (Less Reliable)

Attempt download using other methods:

```bash
python3 download_gilts_data.py  # Using requests
```

```bash
python3 download_gilts_data_advanced.py  # Using requests-html
```

### Processing Downloaded Data

Process any downloaded Gilts data file:

```bash
python3 process_gilts_data.py /path/to/downloaded/file.xlsx [date_in_DD/MM/YYYY]
```

For example:

```bash
python3 process_gilts_data.py ~/Downloads/D1A.xlsx 19/03/2025
```

To check the content of an Excel file (especially useful for .xls format):

```bash
python3 check_excel.py /path/to/downloaded/file.xlsx
```

The processing script will:

1. Copy the file to the `downloads` folder with proper naming
2. Parse and display basic information about the data
3. Offer to save the data as CSV for further analysis

### Manual Download (If Needed)

If you prefer to download manually:

1. Visit the [DMO website](https://www.dmo.gov.uk/data/gilt-market/gilts-in-issue/)
2. Accept any cookie popups
3. Click the 'Excel' button to download the file
4. Process the downloaded file using the instructions above

## Data Source

Data is sourced from the UK Debt Management Office:
[https://www.dmo.gov.uk/data/gilt-market/gilts-in-issue/](https://www.dmo.gov.uk/data/gilt-market/gilts-in-issue/)

## File Format

The downloaded file is in the older Excel (.xls) format and contains information about UK Gilts in Issue, including:

- ISIN codes
- Gilt names
- Coupon rates
- Maturity dates
- Outstanding amounts

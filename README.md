# Job-Scraper-Project

## Overview
This project includes a Python script that scrapes job listings from the Bentley Systems careers page and saves them to an Excel file. The Excel file is styled to match a specific design.

## Files
- `import_requests.py`: The Python script that performs the web scraping and Excel file generation.
- `bentley_jobs.xlsx`: The generated Excel file with job listings.
- `README.md`: This file with instructions and project details.

## Requirements
- Python 3.12+
- Required Python libraries: `requests`, `lxml`, `pandas`, `xlsxwriter`

## How to Run the Script
1. **Install the required libraries**:
   ```sh
   pip install requests lxml pandas xlsxwriter
2. Run the script:
    py import_requests.py
3. Output:
     The script will generate an Excel file named bentley_jobs.xlsx in the same directory.

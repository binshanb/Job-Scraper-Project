import requests
from lxml import html
import pandas as pd
import time

# Base URL of the Bentley Systems careers page
base_url = "https://jobs.bentley.com/search"

# Function to get job listings from a specific page
def get_job_listings(start_row):
    params = {
        'q': '',
        'sortColumn': 'sort_title',
        'sortDirection': 'desc',
        'startrow': start_row
    }
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = requests.get(base_url, params=params)
            response.raise_for_status()  # Ensure we notice bad responses
            tree = html.fromstring(response.content)
            
            # Extract job details using updated XPath expressions
            job_titles = tree.xpath('//tr[contains(@class, "data-row")]/td[contains(@class, "colTitle")]/span/a[contains(@class, "jobTitle-link")]/text()')
            job_locations = tree.xpath('//span[contains(@class, "jobLocation")]/text()')
            job_dates = tree.xpath('//span[contains(@class, "jobDate")]/text()')
            
            # Debug print statements to verify data extraction
            for title, location, date in zip(job_titles, job_locations, job_dates):
                print("Job Title:", title.strip())
                print("Location:", location.strip())
                print("Date Posted:", date.strip())
                print("-------------------------")
            
            jobs = [
                {'Job Title': title.strip(), 'Location': location.strip(), 'Date Posted': date.strip()}
                for title, location, date in zip(job_titles, job_locations, job_dates)
            ]
            return jobs
        
        except requests.exceptions.RequestException as e:
            print(f"Request failed (attempt {attempt + 1}): {e}")
            time.sleep(5)  # Wait a bit before retrying
    
    return []

# Initialize an empty list to store all job listings
all_jobs = []

# Loop through multiple pages (example: first 5 pages)
start_row = 0
while True:
    jobs = get_job_listings(start_row)
    if not jobs:
        break  # Stop if no more jobs are found
    all_jobs.extend(jobs)
    start_row += 25  # Assuming each page has 25 job listings
    
    # Check if there are more jobs to fetch
    response = requests.get(base_url, params={'startrow': start_row})
    tree = html.fromstring(response.content)
    if not tree.xpath('//tr[contains(@class, "datarow")]'):
        break  # Stop if no more jobs are found

# Convert the list of job dictionaries to a DataFrame
df = pd.DataFrame(all_jobs)

# Create an Excel writer object using xlsxwriter
writer = pd.ExcelWriter('bentley_jobs.xlsx', engine='xlsxwriter')

# Write DataFrame to Excel sheet
df.to_excel(writer, sheet_name='Job Listings', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook  = writer.book
worksheet = writer.sheets['Job Listings']

# Add some basic formatting
header_format = workbook.add_format({
    'bold': True, 
    'border': 1, 
    'bg_color': '#F0F0F0', 
    'align': 'center'
})
title_format = workbook.add_format({
    'bold': True,
    'border': 1, 
    'bg_color': '#E8E8E8', 
    'align': 'left',
    'font_color': '#0070C0'
})
location_format = workbook.add_format({
    'border': 1, 
    'bg_color': '#F8F8F8', 
    'align': 'left',
    'font_color': '#000000'
})
date_format = workbook.add_format({
    'border': 1, 
    'bg_color': '#F8F8F8', 
    'align': 'right',
    'font_color': '#000000'
})

# Apply formatting to headers
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)

# Apply alternating row formatting
for row_num, row_data in enumerate(df.values, start=1):
    row_format = workbook.add_format({'bg_color': '#F0F0F0'}) if row_num % 2 == 0 else workbook.add_format({'bg_color': '#FFFFFF'})
    
    worksheet.write(row_num, 0, row_data[0], title_format)
    worksheet.write(row_num, 1, row_data[1], location_format)
    worksheet.write(row_num, 2, row_data[2], date_format)
    worksheet.set_row(row_num, None, row_format)

# Adjust column widths
worksheet.set_column('A:A', 40)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 15)

# Hide all gridlines
worksheet.hide_gridlines(option=2)

# Close the Excel writer
writer._save()

print("Job listings have been saved to bentley_jobs.xlsx")



























# import requests
# from lxml import html
# import pandas as pd

# # Base URL of the Bentley Systems careers page
# base_url = "https://jobs.bentley.com/search"

# # Function to get job listings from a specific page
# def get_job_listings(start_row):
#     params = {
#         'q': '',
#         'sortColumn': 'sort_title',
#         'sortDirection': 'desc',
#         'startrow': start_row
#     }
#     response = requests.get(base_url, params=params)
#     response.raise_for_status()  # Ensure we notice bad responses
#     tree = html.fromstring(response.content)
    
#     # Extract job details using updated XPath expressions
#     job_titles = tree.xpath('//tr[contains(@class, "data-row")]/td[contains(@class, "colTitle")]/span/a[contains(@class, "jobTitle-link")]/text()')
#     job_locations = tree.xpath('//span[contains(@class, "jobLocation")]/text()')
#     job_dates = tree.xpath('//span[contains(@class, "jobDate")]/text()')

    
#     # Debug print statements to verify data extraction
#     for title, location, date in zip(job_titles, job_locations, job_dates):
#       print("Job Title:", title.strip())
#       print("Location:", location.strip())
#       print("Date Posted:", date.strip())
#       print("-------------------------")
    
#     jobs = [
#         {'Job Title': title.strip(), 'Location': location.strip(), 'Date Posted': date.strip()}
#         for title, location, date in zip(job_titles, job_locations, job_dates)
#     ]
#     return jobs

# # Initialize an empty list to store all job listings
# all_jobs = []

# # Loop through multiple pages (example: first 5 pages)
# start_row = 0
# while True:
#     jobs = get_job_listings(start_row)
#     if not jobs:
#         break  # Stop if no more jobs are found
#     all_jobs.extend(jobs)
#     start_row += 25  # Assuming each page has 25 job listings
#     # Check if there are more jobs to fetch
#     response = requests.get(base_url, params={'startrow': start_row})
#     tree = html.fromstring(response.content)
#     if not tree.xpath('//tr[contains(@class, "datarow")]'):
#         break  # Stop if no more jobs are found
# # Convert the list of job dictionaries to a DataFrame
# df = pd.DataFrame(all_jobs)
# # Create an Excel writer object using xlsxwriter
# writer = pd.ExcelWriter('bentley_jobs.xlsx', engine='xlsxwriter')

# # Write DataFrame to Excel sheet
# df.to_excel(writer, sheet_name='Job Listings', index=False)

# # Get the xlsxwriter workbook and worksheet objects
# workbook  = writer.book
# worksheet = writer.sheets['Job Listings']

# # Add some basic formatting
# header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F0F0F0'})
# data_format = workbook.add_format({'border': 1})

# # Apply formatting to headers and data
# for col_num, value in enumerate(df.columns.values):
#     worksheet.write(0, col_num, value, header_format)
    
# for row_num, row_data in enumerate(df.values, start=1):
#     for col_num, value in enumerate(row_data):
#         worksheet.write(row_num, col_num, value, data_format)

# # Close the Excel writer
# writer._save()

# print("Job listings have been saved to bentley_jobs.xlsx")

 









# import requests
# import pandas as pd
# from lxml import html
# import re

# def fetch_data(url):
#     """
#     Fetches HTML content from the specified URL using the requests module.
    
#     Args:
#     - url (str): The URL to fetch data from.
    
#     Returns:
#     - str: The HTML content retrieved from the URL.
#     """
#     try:
#         response = requests.get(url)
#         response.raise_for_status()  # Raise an error for non-200 status codes
#         return response.text
#     except requests.RequestException as e:
#         print("Error fetching data:", e)
#         return None

# def extract_data(html_content):
#     """
#     Extracts job data from HTML content using XPath and regular expressions.
    
#     Args:
#     - html_content (str): The HTML content to extract data from.
    
#     Returns:
#     - tuple: A tuple containing lists of job titles, locations, links, and dates.
#     """
#     if html_content:
#         tree = html.fromstring(html_content)
        
#         # XPath queries to extract job data
#         job_titles = tree.xpath("//div[@class='job-title']/text()")
#         job_locations = tree.xpath("//span[@class='location']/text()")
#         job_links = tree.xpath("//a[@class='job-link']/@href")
#         job_dates = tree.xpath("//span[@class='job-date']/text()")
        
#         return job_titles, job_locations, job_links, job_dates
#     else:
#         return [], [], [], []

# def process_data(job_titles, job_locations, job_links, job_dates):
#     """
#     Processes the extracted job data.
    
#     Args:
#     - job_titles (list): List of job titles.
#     - job_locations (list): List of job locations.
#     - job_links (list): List of job links.
#     - job_dates (list): List of job dates.
    
#     Returns:
#     - list: A list of dictionaries containing processed job data.
#     """
#     # Process the extracted data as needed
#     # Example: Combine the data into a list of dictionaries
#     data = [{'Job Title': title, 'Location': location, 'Link': link, 'Date': date}
#             for title, location, link, date in zip(job_titles, job_locations, job_links, job_dates)]
#     return data

# def save_to_excel(data, filename):
#     """
#     Saves job data to an Excel file.
    
#     Args:
#     - data (list): List of dictionaries containing job data.
#     - filename (str): The name of the Excel file to save.
#     """
#     try:
#         df = pd.DataFrame(data)
#         df.to_excel(filename, index=False)
#         print("Data saved to", filename)
#     except Exception as e:
#         print("Error saving data to Excel:", e)

# def main():
#     """
#     Main function to orchestrate the scraping process.
#     """
#     base_url = 'https://jobs.bentley.com/job/'
#     all_jobs_data = []

#     page = 1
#     while True:
#         url = base_url.format(page)
#         html_content = fetch_data(url)
#         job_titles, job_locations, job_links, job_dates = extract_data(html_content)
        
#         if not job_titles:
#             break  # No more jobs on this page
        
#         data = process_data(job_titles, job_locations, job_links, job_dates)
#         all_jobs_data.extend(data)
        
#         page += 25
#     save_to_excel(all_jobs_data, 'jobs.xlsx')

# if __name__ == "__main__":
#     main()
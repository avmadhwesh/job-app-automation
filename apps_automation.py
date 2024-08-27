import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from datetime import datetime
import re

file_path = 'Applications_Fall.xlsx'
workbook = load_workbook(file_path)
sheet = workbook.active


def get_job_description(job_link):
    response = requests.get(job_link)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find all div elements that have "description" in their class or id
    description_divs = soup.find_all('div', attrs={'class': re.compile(r'description', re.I), 'id': re.compile(r'description', re.I)})

    # Collect text from these divs
    description = "\n".join([div.get_text(strip=True) for div in description_divs])

    # If no divs with "description" are found, get all the text on the page
    if not description:
        full_text = soup.get_text(separator=' ', strip=True)
        description = full_text

    # Further split the text if "Equal Opportunity Employer" is found
    if "Equal Opportunity Employer" in description:
        description = description.split("Equal Opportunity Employer")[0]

    return description.strip() if description else "No description found"





def add_job_application(job_link, referral):
    job_description = get_job_description(job_link)
    date_applied = datetime.now().strftime('%Y-%m-%d')

    next_row = sheet.max_row + 1

    sheet[f'F{next_row}'] = job_link       
    sheet[f'E{next_row}'] = job_description 
    sheet[f'D{next_row}'] = date_applied     #
    sheet[f'G{next_row}'] = referral

    workbook.save(file_path)
    print(f"Data saved for {job_link}")

# Example usage
job_link = input("Enter the job link: ")
referral = input("Referred? ")
add_job_application(job_link, referral)

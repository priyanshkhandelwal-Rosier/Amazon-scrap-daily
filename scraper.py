import pandas as pd
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from openpyxl import load_workbook
from openpyxl.styles import Font
import os

# 1. HTML File Load (Jo GitHub repo me hogi)
try:
    with open("Amazon.html", "r", encoding="utf-8") as file:
        html_content = file.read()
except FileNotFoundError:
    print("Error: 'Amazon.html' repo me nahi mili.")
    exit()

soup = BeautifulSoup(html_content, 'html.parser')
products_data = []

product_divs = soup.find_all('div', {'data-component-type': 's-search-result'})
print(f"Products Found: {len(product_divs)}")

for item in product_divs:
    # Filter Logic (ROSIER)
    brand_tag = item.find('span', class_='a-size-base-plus a-color-base')
    if brand_tag and "ROSIER" in brand_tag.get_text().strip():
        
        # Name
        name_text = "Unknown"
        h2_tag = item.find('h2', class_='a-text-normal')
        if h2_tag and h2_tag.span:
            name_text = h2_tag.span.get_text().strip()

        # Link Extraction
        product_link = None
        if h2_tag:
            parent_link = h2_tag.find_parent('a')
            if parent_link:
                product_link = "https://www.amazon.in/s?k=rosier+foods&crid=84CJ0Q6WFCI4&sprefix=rosier+foo%2Caps%2C481&ref=nb_sb_noss_2" + parent_link['href']

        # MRP
        mrp_text = "N/A"
        price_tag = item.find('span', class_='a-price-whole')
        if price_tag:
            mrp_text = price_tag.get_text().strip()

        # Stock
        stock_text = "Available"
        if item.find('span', class_='a-color-success'):
            stock_text = item.find('span', class_='a-color-success').get_text().strip()
        elif "Currently unavailable" in item.get_text():
            stock_text = "Out of Stock"

        products_data.append({
            'Brand': 'ROSIER',
            'Product Name': name_text,
            'MRP': mrp_text,
            'Stock': stock_text,
            'Hidden_URL': product_link
        })

# 2. Excel Creation with Fixed Hyperlinks
file_name = "Rosier_Report.xlsx"
if products_data:
    df = pd.DataFrame(products_data)
    df.to_excel(file_name, index=False)
    
    # Fix Links on MRP column
    wb = load_workbook(file_name)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        mrp_cell = row[2] # Column C
        url_cell = row[4] # Column E
        if url_cell.value:
            mrp_cell.hyperlink = url_cell.value
            mrp_cell.font = Font(color="0000FF", underline="single")
    ws.delete_cols(5) # Remove URL column
    wb.save(file_name)
else:
    print("No ROSIER products found in the HTML file.")
    exit()

# 3. Email Sending (Using GitHub Secrets)
SENDER_EMAIL = os.environ.get('EMAIL_USER')
SENDER_PASSWORD = os.environ.get('EMAIL_PASS')
RECEIVER_EMAIL = "ismat.saifi@rosierfoods.com" # Yahan receive karne wale ka email likhein

if not SENDER_EMAIL or not SENDER_PASSWORD:
    print("Error: GitHub Secrets set nahi hain.")
    exit()

msg = MIMEMultipart()
msg['From'] = SENDER_EMAIL
msg['To'] = RECEIVER_EMAIL
msg['Subject'] = "Daily Report: Amazon Scrap Data"
msg.attach(MIMEText("Attached is the processed Excel file from Amazon.html", 'plain'))

with open(file_name, "rb") as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {file_name}")
    msg.attach(part)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(SENDER_EMAIL, SENDER_PASSWORD)
server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
server.quit()
print("Email Sent Successfully!")

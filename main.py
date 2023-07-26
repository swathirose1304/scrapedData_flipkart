import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook

# Function to scrape data and save to a worksheet
def scrape_data(url, worksheet_name):
    product_name_lst = []
    real_price_lst = []
    offer_percent_lst = []
    offer_prices_lst = []

    r = requests.get(url)
    soup = BeautifulSoup(r.text, 'lxml')
    box = soup.find('div', class_='_1YokD2 _3Mn1Gg')

    names = box.find_all('div', class_='_2WkVRV')
    for i in names:
        name = i.text
        product_name_lst.append(name)

    real_prices = box.find_all('div', class_='_3I9_wc')
    for i in real_prices:
        price = i.text
        real_price_lst.append(price)

    offers = box.find_all('div', class_='_3Ay6Sb')
    for i in offers:
        offer = i.text
        offer_percent_lst.append(offer)

    off_prices = box.find_all('div', class_='_30jeq3')
    for i in off_prices:
        price = i.text
        offer_prices_lst.append(price)

    # Fill empty values with empty strings
    max_length = max(len(product_name_lst), len(real_price_lst), len(offer_percent_lst), len(offer_prices_lst))
    product_name_lst += [''] * (max_length - len(product_name_lst))
    real_price_lst += [''] * (max_length - len(real_price_lst))
    offer_percent_lst += [''] * (max_length - len(offer_percent_lst))
    offer_prices_lst += [''] * (max_length - len(offer_prices_lst))

    # Create a worksheet in the existing workbook
    worksheet = workbook.create_sheet(worksheet_name)

    # Convert the scraped data into a DataFrame
    df = pd.DataFrame({
        'ProductName': product_name_lst,
        'RealPrice': real_price_lst,
        'Offer': offer_percent_lst,
        'OfferPrice': offer_prices_lst
    })

    # Write the DataFrame to the worksheet
    header = list(df.columns)  # Get the column headers
    worksheet.append(header)  # Write the header row

    # Write the DataFrame rows to the worksheet
    for _, row in df.iterrows():
        worksheet.append(row.tolist())

# Create a new workbook
workbook = Workbook()

# Scrape data for 'shoes' and save to worksheet
shoes_url = 'https://www.flipkart.com/search?q=shoes&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off&as-pos=1&as-type=HISTORY&page='+str(1)
scrape_data(shoes_url, 'shoes')

# Scrape data for 'watches' and save to worksheet
watches_url = 'https://www.flipkart.com/search?q=watches&otracker=AS_Query_HistoryAutoSuggest_2_0&otracker1=AS_Query_HistoryAutoSuggest_2_0&marketplace=FLIPKART&as-show=on&as=off&as-pos=2&as-type=HISTORY&page='+str(1)
scrape_data(watches_url, 'watches')

# Save the workbook
workbook.save('flipkart_data.xlsx')
print('Saved')

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook

product_name_lst = []
real_price_lst = []
offer_percent_lst = []
offer_prices_lst = []

# for i in range(2,10):
watches_url = 'https://www.flipkart.com/search?q=watches&otracker=AS_Query_HistoryAutoSuggest_2_0&otracker1=AS_Query_HistoryAutoSuggest_2_0&marketplace=FLIPKART&as-show=on&as=off&as-pos=2&as-type=HISTORY&page='+str(1)
r = requests.get(watches_url)
# print(r)

soup = BeautifulSoup(r.text, 'lxml')
# print(soup)
box = soup.find('div', class_='_1YokD2 _3Mn1Gg')   # scrap data from box...not from complete page

names = box.find_all('div', class_='_2WkVRV')

for i in names:
    name = i.text
    product_name_lst.append(name)
# print(product_name_lst)

real_prices = box.find_all('div', class_='_3I9_wc')

for i in real_prices:
    price = i.text
    real_price_lst.append(price)
# print(real_price_lst)

offers = box.find_all('div', class_='_3Ay6Sb')

for i in offers:
    offer = i.text
    offer_percent_lst.append(offer)
# print(offer_percent_lst)

off_prices = box.find_all('div', class_='_30jeq3')

for i in off_prices:
    price = i.text
    offer_prices_lst.append(price)
# print(offer_prices_lst)

# Create a new workbook
workbook = Workbook()

# Create a worksheet named 'shoes'
worksheet = workbook.create_sheet('watches')

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


# Save the workbook
workbook.save('flipkart_data.xlsx')
print('saved')



















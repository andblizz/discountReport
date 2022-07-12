import os
import requests
import base64
import xlsxwriter
from bs4 import BeautifulSoup
import datetime


headers = {'content-type': 'text/xml'}

# Request actions
fromDate = str(input('Введите дату отчёта в формате ГГГГ-ММ-ДД: '))
year, month, day = map(int, fromDate.split('-'))
toDate = str(datetime.date(year, month, day)+datetime.timedelta(1))
try:
    os.mkdir('Temp_files')
except FileExistsError:
    pass
except:
    print('Unable to create "Temp files" folder')

body_action = f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:plug="http://plugins.operday.ERPIntegration.crystals.ru/">
   <soapenv:Header/>
   <soapenv:Body>
      <plug:getLoyResultsByPeriod>
         <!--Optional:-->
         <fromDate>{fromDate}</fromDate>
         <!--Optional:-->
         <toDate>{toDate}</toDate>
      </plug:getLoyResultsByPeriod>
   </soapenv:Body>
</soapenv:Envelope>"""
response_action = requests.post(os.environ['URL'], data=body_action, headers=headers)
with open('Temp_files/loyResults.xml', 'wb') as f:
    f.write(response_action.content)

with open('Temp_files/loyResults.xml', 'r') as f:
    soup = BeautifulSoup(f, 'xml')

with open('Temp_files/purchases_action_base64.xml', 'w') as f:
    f.write(soup.find('return').text)

with open('Temp_files/purchases_action_decode.xml', 'wb') as f:
    decoded = base64.b64decode(open('Temp_files/purchases_action_base64.xml', 'r').read())
    f.write(decoded)


# Request all checks
body_all_checks = f"""<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:plug="http://plugins.operday.ERPIntegration.crystals.ru/">
   <soapenv:Header/>
   <soapenv:Body>
      <plug:getPurchasesByOperDay>
         <!--Optional:-->
         <dateOperDay>{fromDate}</dateOperDay>
      </plug:getPurchasesByOperDay>
   </soapenv:Body>
</soapenv:Envelope>"""

response_all_checks = requests.post(os.environ['URL'], data=body_all_checks, headers=headers)
purchase_base64 = open('Temp_files/purchase_all_checks.xml', 'wb')
purchase_base64.write(response_all_checks.content)
purchase_base64.close()

with open('Temp_files/purchase_all_checks.xml', 'r') as f:
    soup = BeautifulSoup(f, 'xml')

with open('Temp_files/purchase_all_checks_base64.xml', 'w') as f:
    f.write(soup.find('return').text)

with open('Temp_files/purchase_all_checks_decode.xml', 'wb') as f:
    decoded = base64.b64decode(open('Temp_files/purchase_all_checks_base64.xml', 'r').read())
    f.write(decoded)


with open('Temp_files/purchases_action_decode.xml') as f:
    soup = BeautifulSoup(f, 'xml')
purchase = soup.find_all('purchase')
purchases = soup.find_all('purchases')
purchases_count = int(purchases[0]['count'])

with open('Temp_files/purchase_all_checks_decode.xml', encoding='utf-8') as f:
    soup = BeautifulSoup(f, 'xml')
purchases_all_checks = soup.find_all('purchases')
purchases_all_checks_count = int(purchases_all_checks[0]["count"])
purchase_all_checks = soup.find_all('purchase')


action_id = "82397524"
amounts = []
card_numbers = []
discountValueTotals = []
shops = []
saletimes_action = []
check_amounts = []
saletimes_all_checks = []
for i in purchase_all_checks:
    saletimes_all_checks.append(i['saletime'][:23])


def counter(attribute):
    if attribute == 'card_number':
        for i in range(purchases_count):
            try:
                if purchase[i].discount['AdvertActGUID'] == action_id:
                    card_numbers.append(purchase[i].text.strip())
            except:
                pass
    if attribute == 'discountValueTotal':
        for i in range(purchases_count):
            try:
                if purchase[i].discount['AdvertActGUID'] == action_id:
                    discountValueTotals.append(purchase[i].discount.parent['discountValueTotal'])
            except:
                pass
    if attribute == 'shop':
        for i in range(purchases_count):
            try:
                if purchase[i].discount['AdvertActGUID'] == action_id:
                    shops.append(purchase[i].discount.parent['shop'])
            except:
                pass
    if attribute == 'saletime':
        for i in range(purchases_count):
            try:
                if purchase[i].discount['AdvertActGUID'] == action_id:
                    saletimes_action.append(purchase[i].discount.parent['saletime'])
            except:
                pass


counter('discountValueTotal')
counter('card_number')
counter('shop')
counter('saletime')


saletime_check_amount_dict = {}
for i in range(purchases_all_checks_count):
    for io in saletimes_action:
        if io in purchase_all_checks[i]['saletime'][:23]:
            saletime_check_amount_dict.update({purchase_all_checks[i]['saletime'][:23]: purchase_all_checks[i]['amount']})


for element in saletimes_action:
    check_amounts.append(saletime_check_amount_dict.get(element))


try:
    os.mkdir('Reports')
except FileExistsError:
    pass
except:
    print('Unable to create "Reports" folder')
workbook = xlsxwriter.Workbook(f'Reports/{saletimes_action[0][:10]}.xlsx')
worksheet = workbook.add_worksheet()

row_card = 1
row_shop = 1
row_saletime = 1
row_check_amount = 1
row_discountValueTotal = 1
col = 0

headings_format = workbook.add_format({'bold': True, 'border': 2})
headings = ['Shop', 'Card number', 'Sale time', 'Check amount', 'Discount value total']
worksheet.write_row('A1', headings, headings_format)


for shop in shops:
    worksheet.write(row_shop, col, shop)
    row_shop += 1
for card in card_numbers:
    worksheet.write(row_card, col + 1, str(card.split('\n')))
    row_card += 1
for saletime in saletimes_action:
    worksheet.write(row_saletime, col + 2, saletime)
    row_saletime += 1
for check_amount in check_amounts:
    worksheet.write(row_check_amount, col + 3, check_amount)
    row_check_amount += 1
for dVT in discountValueTotals:
    worksheet.write(row_discountValueTotal, col + 4, dVT)
    row_discountValueTotal += 1
workbook.close()

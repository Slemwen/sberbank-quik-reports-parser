import os
import glob
import pandas as pd
from bs4 import BeautifulSoup
import requests
import datetime
from math import isinf

# files directory
directory = os.getcwd()

for filename in glob.glob(os.path.join(directory, '*.xlsx')):
	if filename != 'report.xlsx':
		data = pd.read_excel(filename)
# creating a start table
table_xlsx = data.loc[data['Срок расчётов'] == 'T2'][['Название инструмента',
                                                      'Код инструмента', 'Баланс', 'Цена приобретения']]
table_xlsx = table_xlsx.reset_index(drop=True)
table_xlsx['Цена приобретения'] = table_xlsx['Цена приобретения'].round(2)

#  collecting new data
data_from_html = {}
for i in range(len(table_xlsx)):
	url = 'https://smart-lab.ru/q/bonds/' + table_xlsx['Код инструмента'][i]
	#  finding necessary data in html
	lab = requests.get(url)
	html_soup = BeautifulSoup(lab.content.decode(lab.encoding), 'html.parser')
	text = html_soup.text.split('\n')
	redemption_name = next(x for x in text if 'Дата погашения' in x)
	data_from_html.setdefault('Дата погашения', []).append(text[text.index(redemption_name) + 1])

	aci_ncd = next(x for x in text if 'НКД' in x)
	data_from_html.setdefault('НКД', []).append(text[text.index(aci_ncd) + 1].split()[0])

	value = next(x for x in text if 'Номинал' in x)
	data_from_html.setdefault('Номинал', []).append(text[text.index(value) + 1])

	current_price = next(x for x in text if 'Цена послед' in x)
	data_from_html.setdefault('Текущая цена', []).append(text[text.index(current_price) + 1])

	cupon_percent = next(x for x in text if 'Дох. купона' in x)
	data_from_html.setdefault('Доходность купона, %', []).append(text[text.index(cupon_percent) + 1][:-1])

	cupon_rubles = next(x for x in text if 'Купон, руб' in x)
	data_from_html.setdefault('Купон, руб', []).append(text[text.index(cupon_rubles) + 1].split()[0])

	cupon_period = next(x for x in text if 'Выплата купона' in x)
	data_from_html.setdefault('Выплата купона', []).append(text[text.index(cupon_period) + 1])

	data_from_html.setdefault('Ссылка на форум', []).append(url)

#  adding new data
table_xlsx['Дата погашения'] = data_from_html['Дата погашения']
table_xlsx['НКД'] = data_from_html['НКД']
table_xlsx['Номинал'] = data_from_html['Номинал']
table_xlsx['Текущая цена'] = data_from_html['Текущая цена']
table_xlsx['Доходность купона, %'] = data_from_html['Доходность купона, %']
table_xlsx['Купон, руб'] = data_from_html['Купон, руб']
table_xlsx['Выплата купона, дней'] = data_from_html['Выплата купона']
table_xlsx['Ссылка на форум'] = data_from_html['Ссылка на форум']


#  calculate new data
table_xlsx = table_xlsx.astype(
	{'Номинал': 'float', 'Баланс': 'float', 'НКД': 'float', 'Текущая цена': 'float', 'Цена приобретения': 'float', 'Купон, руб': 'float',
	 'Выплата купона, дней': 'float', 'Доходность купона, %': 'float'})
table_xlsx['Стоимость позиции'] = (table_xlsx['Баланс'] * table_xlsx['Текущая цена'] * (
		table_xlsx['Номинал'] + table_xlsx['НКД']) / 100).round()

#  import from second report file
for html_report in glob.glob(os.path.join(directory, '*.htm')):
	with open(html_report) as report:
		html_soup_money = BeautifulSoup(report, 'html.parser')
#  find the cash
table_money = html_soup_money.find_all("table")
cash_string = list(table_money[-2])[-2]
data_money = float([th.get_text() for th in cash_string.find_all("td")][-1].replace(' ', ''))
table_xlsx = table_xlsx.append({'Название инструмента': 'Денежные средства'}, ignore_index=True)
table_xlsx.iloc[-1, table_xlsx.columns.get_loc('Стоимость позиции')] = round(data_money)
#  sum of assets
sum_of_assets = sum(table_xlsx['Стоимость позиции'])
table_xlsx = table_xlsx.append({'Название инструмента': 'Сумма активов'}, ignore_index=True)
table_xlsx.iloc[-1, table_xlsx.columns.get_loc('Стоимость позиции')] = sum_of_assets
#  % of assets
table_xlsx['% от активов'] = (100 * table_xlsx['Стоимость позиции'] / sum_of_assets).round(1)
#  time to redemption of bonds
today = datetime.date.today()
years_to_redemption = []
for i in table_xlsx['Дата погашения']:
	if not pd.isna(i):
		years_to_redemption.append(
			round((datetime.date.fromisoformat('-'.join(str(i).split('-')[::-1])) - today).days / 365, 1))
	else:
		years_to_redemption.append(0)
table_xlsx['Лет до погашения'] = years_to_redemption

#  profitability calculation
table_xlsx['Моя доходность купона'] = (((table_xlsx['Номинал'] - table_xlsx['Цена приобретения'] * table_xlsx['Номинал'] / 100) + table_xlsx['Купон, руб'] * 365 / table_xlsx['Выплата купона, дней'] * table_xlsx['Лет до погашения']) / (table_xlsx['Цена приобретения'] * table_xlsx['Номинал'] / 100) * 100 / table_xlsx['Лет до погашения']).round(1)

table_xlsx['Результат, %'] = (100 * (table_xlsx['Текущая цена'] - table_xlsx['Цена приобретения']) / table_xlsx['Цена приобретения']).round(1)

data_coupon = table_xlsx['Моя доходность купона'] * table_xlsx['Баланс']

data_coupon_sum = 0
for i in data_coupon:
	if not pd.isna(i) and not isinf(i):
		data_coupon_sum += i

data_balance_sum = 0
for i in table_xlsx['Баланс']:
	if not pd.isna(i) and not isinf(i):
		data_balance_sum += i
table_xlsx = table_xlsx.append({'Название инструмента': 'Доходность портфеля'}, ignore_index=True)
table_xlsx.iloc[-1, table_xlsx.columns.get_loc('Моя доходность купона')] = round(data_coupon_sum / data_balance_sum, 2)

#  wriring to excel

table_xlsx.to_excel('report.xlsx')
# print(table_xlsx['Доходность купона, %'])

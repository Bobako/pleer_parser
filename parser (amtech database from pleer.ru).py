import requests


from bs4 import BeautifulSoup


from file_processer import get_parts

from selenium import webdriver 
from time import sleep


driver = webdriver.Firefox(executable_path="E:\webdriver\geckodriver.exe")

	

def get_html(url,driver = driver):

	driver.get(url)
	while driver.title == '':
		sleep(0.1)
	
	html = driver.page_source
	return html

def get_price(html):
	soup = BeautifulSoup(html, 'lxml')
	price = soup.find_all('div', class_ = 'product_price_div')[0].find('div',class_='price').text

	price = price.split(' ')[:-1:]

	s = ''
	for p in price:
		s+=str(p)

	s = int(s)

	return s


def get_price_from_url(url):
	html = get_html(url)
	try:
		return get_price(html)
	except Exception:
		return None	


from file_processer import add_row_xlsx

def main():
	parts = get_parts('db_v2.xlsx','Sheet1')
	new_parts = []
	for part in parts:
		if part['name']==None:
			break
		new_price = part['price']
		link = part['ssylka']

		if link!=None: n = get_price_from_url(link)
		if n!= None: new_price = n
		print(part['id'],' - ',n)
		temp = part
		temp['price'] = new_price
		for key in temp:
			try:
				temp[key] = int(temp[key])
			except Exception:
				pass
				
			if temp[key]=='None' or temp[key]==None:
				temp[key]=''
                
		add_row_xlsx('db_v2.1.xlsx','Sheet1',temp)

	return new_parts











main()

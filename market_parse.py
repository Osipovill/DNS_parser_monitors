import time
import re
from fake_useragent import UserAgent
from selenium.webdriver import ActionChains
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
from openpyxl.drawing.image import Image

# вводим параметры парсера
options = Options()
ua = UserAgent()
userAgent = ua.random
options.add_argument(f'user-agent={userAgent}')
options.add_argument('--disable-blink-features=AutomationControlled')
service = Service('C:\\parsing\\msedgedriver.exe')
driver = webdriver.Edge(service=service, options=options)

# вводим параметры эксель
wb = openpyxl.Workbook()
ws = wb.active
titles = ["изображение","ссылка на товар","название",  "цена", "рейтинг","кол-во отзывов"]
ws.append(titles)

# парсим страницу
def parse(item):

    res = []
    image_url = item.find_element(by=By.TAG_NAME, value="source").get_attribute('data-srcset')
    url = item.find_element(by=By.CLASS_NAME, value="catalog-product__image-link").get_attribute('href')
    name = item.find_element(by=By.CLASS_NAME, value="catalog-product__name.ui-link.ui-link_black").text
    cost = item.find_element(by=By.CLASS_NAME, value="product-buy__price").text
    feedback = item.find_element(by=By.CLASS_NAME, value="catalog-product__rating.ui-link.ui-link_black").get_attribute('data-rating')
    feedback_count = item.find_element(by=By.CLASS_NAME, value="catalog-product__rating.ui-link.ui-link_black").text
    res.append(image_url)
    res.append(url)
    res.append(name)
    res.append(cost)
    res.append(feedback)
    res.append(feedback_count)
    return res


######## Основной код
for page_number in range(1,9):

    url = 'https://www.dns-shop.ru/catalog/17a8943716404e77/monitory/?p=' + str(page_number)
    driver.get(url=url)
    time.sleep(1)
    items = driver.find_elements(by=By.CLASS_NAME, value="catalog-product.ui-button-widget ")
    for item in items:
        data = parse(item)
        ws.append(data)

wb.save("inf.xlsx")
driver.close()

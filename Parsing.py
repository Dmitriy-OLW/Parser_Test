import openpyxl
import requests
import time

from bs4 import BeautifulSoup as bs
from selenium import common
from selenium.webdriver.common.by import By
from selenium import webdriver


def parse(el):
    base = []
    text = (el.find_element(by=By.CLASS_NAME, value="link-wrap").find_elements(
        by=By.TAG_NAME, value="a"))
    base.append(f"{text[0].text}")

    url_p = el.find_element(by=By.CLASS_NAME,
                                value="goods_img_link").get_attribute('href')
    base.append(url_p)

    text = (el.find_element(by=By.CLASS_NAME, value="description").find_elements(
        by=By.TAG_NAME, value="span"))
    if(len(text) == 2):
        base.append(f"{text[0].text}")
        base.append(f"{text[1].text}")
    else:
        base.append(f"{text[0].text}")
        base.append(' ')

    price = el.find_element(by=By.CLASS_NAME, value="price").text
    base.append(price)

    try:
        old_price = el.find_element(by=By.CLASS_NAME,
                                   value="old-price").text
    except (common.NoSuchElementException):
        old_price = ' '
    base.append(old_price)

    request = requests.get(url_p)
    soup = bs(request.text, 'html.parser')
    steps = soup.findAll("p", class_=["ts-otziv"])
    for step in steps:
        date = (step.text).replace('\n', ' ')
        base.append(date)

    return base


def get_olipics(driver):
    scroll(5, 1, driver)
    site = (driver.find_elements(by=By.CLASS_NAME,
                                     value="hover-block "))
    return site


def scroll(count, delay, driver):
    for i in range(count):
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(delay)


def main():
    url = 'https://airsoft-rus.ru/new-product/?page_count=96'
    service = webdriver.ChromeService(executable_path='chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    driver.get(url)
    site = get_olipics(driver)
    wb = openpyxl.Workbook()
    ws = wb.active
    head = ["Наименование товара", "Ссылка на товар", "Артикул", "Производитель", "Цена", "Цена до скидки", "Отзыв на товар", "Отзыв на магазин"]
    ws.append(head)
    for el in site:
        data = parse(el)
        ws.append(data)
    wb.save("parsing_result.xlsx")


if __name__ == "__main__":
    main()











#---------------------------------------------------------------------------

"""import openpyxl
import requests
import time

from bs4 import BeautifulSoup as bs
from selenium import common
from selenium.webdriver.common.by import By
from selenium import webdriver


def parse(el):
    base = []
    text = (el.find_element(by=By.CLASS_NAME, value="o-info").find_elements(
        by=By.TAG_NAME, value="a"))
    base.append(f"{text[0].text}")
    base.append(f"{text[1].text}")
    url_olimp = el.find_element(by=By.CLASS_NAME,
                                value="none_a.black").get_attribute('href')
    base.append(url_olimp)
    raiting = el.find_element(by=By.CLASS_NAME, value="pl_rating").text
    base.append(raiting)

    try:
        dop_info = el.find_element(by=By.CLASS_NAME,
                                   value="none_a.black.olimp_desc").text
    except (common.NoSuchElementException):
        dop_info = ''
    base.append(dop_info)

    request = requests.get(url_olimp)
    soup = bs(request.text, 'html.parser')
    steps = soup.find_all('tr', {'class': "notgreyclass"})
    for step in steps:
        date = (step.text).replace('\n', ' ')
        base.append(date)

    return base


def get_olipics(driver):
    scroll(20, 1, driver)
    olimpics = (driver.find_elements(by=By.CLASS_NAME,
                                     value="fav_olimp.olimpiada "))
    return olimpics


def scroll(count, delay, driver):
    for i in range(count):
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(delay)


def main():
    url = 'https://olimpiada.ru/activities?subject%5B15%5D=on&class=any&type=any&period_date=&period=year'
    service = webdriver.ChromeService(executable_path='chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    driver.get(url)
    olimpics = get_olipics(driver)
    wb = openpyxl.Workbook()
    ws = wb.active
    for el in olimpics:
        data = parse(el)
        ws.append(data)
    wb.save("parsing_result.xlsx")


if __name__ == "__main__":
    main()
"""
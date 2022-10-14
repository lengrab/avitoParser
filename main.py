from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import xlsxwriter

URL = 'https://www.avito.ru/severodvinsk/telefony?cd=1&q=iphone+12'
BASE_URL = 'https://www.avito.ru'
PAUSE_DURATION_SECONDS = 1


class MyXLSXWriter:
    def __init__(self, file_name: str = "demo"):
        self.file_name = file_name

    def write_item_list(self, item_list):
        workbook = xlsxwriter.Workbook(f'{self.file_name}.xlsx')
        worksheet = workbook.add_worksheet()
        top_shift = 1
        left_shift = 1

        worksheet.write(top_shift - 1, left_shift, "Title")
        worksheet.write(top_shift - 1, left_shift + 1, "Price")
        worksheet.write(top_shift - 1, left_shift + 2, "Link")

        for i in range(len(item_list)):
            worksheet.write(i + top_shift, left_shift, item_list[i].title)
            worksheet.write(i + top_shift, left_shift + 1, item_list[i].price)
            worksheet.write(i + top_shift, left_shift + 2, item_list[i].link)

        workbook.close()


class Item:
    def __init__(self, title: str, price: int, link: str):
        self.title = title
        self.price = price
        self.link = link

    def __str__(self):
        return f'{self.title} \nцена: {self.price} \nlink: {self.link}'


def main():
    driver.get(URL)
    sleep(PAUSE_DURATION_SECONDS)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    item_list = []
    excel_writer = MyXLSXWriter("iphone12")

    all_item = soup.findAll('div', {'data-marker': 'item'})
    for item in all_item:
        price = item.find('meta', {'itemprop': 'price'})['content']
        title = item.find('h3', {'itemprop': 'name'}).text
        link = item.find('a')['href']
        item_list.append(Item(title, price, BASE_URL + link))

    excel_writer.write_item_list(item_list)


if __name__ == '__main__':
    try:
        service = Service(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        main()
    except Exception as e:
        print(e)
    finally:
        driver.quit()

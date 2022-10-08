from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

URL = 'https://www.avito.ru/severodvinsk/telefony?cd=1&q=iphone+12'
PAUSE_DURATION_SECONDS = 1

def main():
    driver.get(URL)
    sleep(PAUSE_DURATION_SECONDS)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    
    allItem = soup.findAll('div',  {'data-marker':'item'})
    for item in allItem:
        price = item.find('meta', {'itemprop':'price'})['content']
        title = item.find('h3', {'itemprop':'name'}).text
        link = item.find('a')['href']
        print(f'{title}, \nцена: {price}, \nlink: {link}')
        print()


if __name__ == '__main__':
    try:
        service = Service(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        main()
    except Exception as e:
        print(e)
    finally:
        driver.quit()



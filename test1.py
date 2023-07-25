from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as BS
from selenium.webdriver.common.action_chains import ActionChains
import requests
from openpyxl import load_workbook

def get_data(url):
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    s = Service(executable_path=r"D:\PyCharm\Projekts\Alibaba_Pars\chromedriver.exe")
    driver = webdriver.Chrome(service=s, options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument",{
        'source': '''
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol
        '''
        }
    )

    try:
        times = 100
        driver.get(url=url)
        sleep(2)
        bottom = driver.find_element(By.CLASS_NAME, "ui-footer-seo")
        while times > 0:
            print(times)
            action = ActionChains(driver)
            action.move_to_element(to_element=bottom).perform()
            sleep(1)
            times = times - 1

        with open("Ali_page_item.html", 'w', encoding='utf-8') as file:
            file.write(driver.page_source)
            file.close()

    except Exception as ex:
        print("Wrong driver setup")
    finally:
        driver.close()
        driver.quit()

def get_href(file_path):
    with open(file_path, encoding='utf-8') as file:
        number = 0
        file_links = open('Links.txt', 'a', encoding='utf-8')
        src = file.read()
        soup = BS(src, 'lxml')
        link_case = soup.find('div', class_='hugo4-pc-grid hugo4-pc-grid-5 hugo4-pc-grid-list')
        links = link_case.find_all('div', class_='hugo4-pc-grid-item')
        for items in links:
            link = items.find_next('a').get('href')
            file_links.write(link + '\n')
            print(f"link {number} collected")
            number = number + 1
        file_links.close()

def get_link_data():
    file_links = open('Links.txt', 'r', encoding='utf-8')
    check = True
    sub_url = "1"
    tab_file = 'Items_test1.xlsx'
    wb = load_workbook(tab_file)
    tab_list = wb['data']
    number = 0

    while check is True:
        if sub_url != "":
            try:
                sub_url = file_links.readline()
                response = requests.get(sub_url)
                soup = BS(response.text, 'lxml')
                item_name = soup.find('div', class_='product-title')
                item_price = soup.find('div', class_='price-list').find('div', class_='price')
                item_categories = soup.find('div', class_='detail-next-breadcrumb')
                tab_list.append([item_name.text, item_price.text])
                wb.save(tab_file)
                print(f"Item {number} collected")
                number = number + 1
            except Exception as ex:
                print("Skip")
        else:
            check = False
    wb.save(tab_file)
    wb.close()
    file_links.close()
    print("Collecting data is finished")

def refresh():
    file = open('Categories.txt', 'r', encoding='utf-8')
    file_clean = open('Links.txt', 'w', encoding='utf-8')
    file_clean.write('')
    file_clean.close()
    check = True
    url = "1"
    while check is True:
        if url != "":
            try:
                url = file.readline()
                get_data(url=url)
                get_href('D:\PyCharm\Projekts\Alibaba_Pars\Ali_page_item.html')
            except Exception as ex:
                print(ex)
        else:
            check = False
            file.close()

def main():
    refresh()
    get_link_data()

if __name__ == "__main__":
    main()
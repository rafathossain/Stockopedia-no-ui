import os.path
import time
from os import path
import sys
import xlsxwriter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def create_directory():
    if not path.exists('Downloaded Files'):
        os.mkdir('Downloaded Files')


def login_to_website(url):
    option = webdriver.ChromeOptions()
    option.add_argument("--incognito")

    browser = webdriver.Chrome(options=option)
    browser.get("https://www.stockopedia.com/auth/login/")

    timeout = 30

    try:
        WebDriverWait(browser, timeout).until(
            EC.visibility_of_element_located((By.XPATH, "//input[@class='ui huge fluid blue button']")))
        time.sleep(1)
        browser.find_element_by_id('username').send_keys("sjjhr67@gmail.com")
        browser.find_element_by_id('password').send_keys("Hargreaves")
        browser.find_element_by_name('remember').click()
        time.sleep(1)
        browser.find_element_by_id('auth_submit').click()
        time.sleep(1)
        try:
            WebDriverWait(browser, timeout).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='app__content ng-tns-c186-0']")))
            time.sleep(1)
            browser.get(url)
            time.sleep(1)
            try:
                WebDriverWait(browser, timeout).until(
                    EC.visibility_of_element_located((By.XPATH, "//div[@class='app__content ng-tns-c186-0']")))
                header = browser.find_elements_by_tag_name("th")
                header = [x.text for x in header]
                del header[0]
                del header[(len(header) - 1)]
                pagination = browser.find_element_by_tag_name("pagination")
                pages = pagination.find_elements_by_tag_name("a")
                pages = [x.text for x in pages]
                lastPage = int(pages[len(pages) - 1])
                current_url = browser.current_url
                filename = time.strftime("Stockopedia_%Y%m%d%H%M%S")
                create_directory()
                stockBook = xlsxwriter.Workbook("Downloaded Files/" + filename + ".xlsx")
                stockSheet = stockBook.add_worksheet()
                for key, val in enumerate(header):
                    stockSheet.write(0, key, val)
                rowCounter = 1
                for pageNumber in range(1, lastPage + 1):
                    if (current_url == "https://app.stockopedia.com/screens/create") or (current_url[0:(
                            len(current_url) - 2)] == "https://app.stockopedia.com/screens/create?page"):
                        browser.get(current_url + "?page=" + str(pageNumber))
                    else:
                        browser.get(current_url + "&page=" + str(pageNumber))
                    time.sleep(1)
                    try:
                        WebDriverWait(browser, timeout).until(
                            EC.visibility_of_element_located((By.XPATH, "//tr[@class='ng-star-inserted']")))
                        time.sleep(1)
                        rows = browser.find_elements_by_tag_name("tr")
                        for index, row in enumerate(rows):
                            data = row.find_elements_by_tag_name("td")
                            if len(data) != 0:
                                flagQuery = data[1].find_element_by_tag_name("span")
                                flagQuery = flagQuery.get_attribute('innerHTML')
                                flag = str(flagQuery).split(' ')[2].split('-')
                                flag = flag[len(flag) - 1]
                                data = [x.text for x in data]
                                del data[0]
                                del data[(len(data) - 1)]
                                data[0] = flag
                                for key, val in enumerate(data):
                                    if "," in str(val):
                                        val = str(val).replace(",", "")
                                    stockSheet.write(rowCounter, key, val)
                                rowCounter += 1
                    # break
                    except TimeoutException:
                        print("Timed out waiting for page to load")
                        break
                stockBook.close()
                browser.quit()
            except TimeoutException:
                print("Timed out waiting for page to load")
                browser.quit()
        except TimeoutException:
            print("Timed out waiting for page to load")
            browser.quit()
    except TimeoutException:
        print("Timed out waiting for page to load")
        browser.quit()


if __name__ == '__main__':
    # print(sys.argv[1])
    login_to_website(sys.argv[1])

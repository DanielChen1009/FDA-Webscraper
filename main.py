from ratelimit import limits, sleep_and_retry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import xlwt

TARGET_LINK = "https://www.accessdata.fda.gov/scripts/cber/CFAppsPub/Index.cfm"
RESULT_LINK_XPATH = "//tr[@class='oddrow' or @class='evenrow']//td[1]//a"
FIELD_XPATH = "//table[@class='StandardTable']//tbody//tr//td[normalize-space(text())='{0}:']/following-sibling::td[1]"
SECONDS_PER_REQUEST = 3
LOAD_WAIT_SECONDS = 20

FILTERS = {
    "EstablishmentType": "3",
    "EstablishmentStatus": "ACTIVE",
    "Country": "US",
    "nrecords": "100"
}

FIELDS = {
    "Address": 1,
    "City": 2,
    "State" : 3,
    "Zip" : 4,
    "Country" : 5,
    "FDA Establishment Identifier (FEI)": 6,
    "Phone": 7
}

MAX_ITEMS_PER_PAGE = 200000
MAX_PAGES = 200000

def run(output_filename, output_sheetname, chromedriver_path):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(output_sheetname)
    ws.write(0, 0, "Name")
    for field, col in FIELDS.items():
        ws.write(0, col, field)

    option = webdriver.ChromeOptions()
    browser = webdriver.Chrome(executable_path=chromedriver_path, options=option)
    browser.get(TARGET_LINK)
    option.add_experimental_option("detach", True)

    wait_load(browser)
    for element_id, value in FILTERS.items():
        select = Select(browser.find_element_by_id(element_id))
        select.select_by_value(value)

    submit_button = browser.find_element_by_id("SubmitButton")
    submit_button.click()
    wait_load(browser)
    row = 1
    pages = 0

    # last = browser.find_element_by_id("Display last")
    # last.click()

    while True:
        links = browser.find_elements_by_xpath(RESULT_LINK_XPATH)
        num = len(links)
        for idx in range(min(MAX_ITEMS_PER_PAGE, num)):
            extract(row, links[idx], browser, ws, wb, output_filename)
            links = browser.find_elements_by_xpath(RESULT_LINK_XPATH)
            row += 1

        if not(check_exists_by_id("Display next", browser)):
            break

        next = browser.find_element_by_id("Display next")
        next.click()
        pages += 1
        if pages >= MAX_PAGES:
            break
    wb.save(output_filename)


def check_exists_by_id(id, browser):
    try:
        browser.find_element_by_id(id)
    except NoSuchElementException:
        return False
    return True


@sleep_and_retry
@limits(calls=1, period=SECONDS_PER_REQUEST)
def extract(row, link, browser, ws, wb, output_filename):
    ws.write(row, 0, link.text)
    wb.save(output_filename)
    link.click()
    wait_load(browser)
    for field, col in FIELDS.items():
        cell = browser.find_element_by_xpath(FIELD_XPATH.format(field))
        ws.write(row, col, cell.text)
    browser.back()


def wait_load(browser):
    try:
        WebDriverWait(browser, LOAD_WAIT_SECONDS).until(EC.visibility_of_element_located((By.ID, "SubmitButton")))
    except TimeoutException:
        print("Timed out loading page")
        browser.quit()


if __name__ == '__main__':
    run('output.xls', 'FDA Plasmapheresis Centers', './chromedriver')

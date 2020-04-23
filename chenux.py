import selenium
import xlrd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui

loc = ("/home/manthila/Downloads/Springer Ebooks.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

linktlist = sheet.col_values(4)

sheet.cell_value(0, 0)




print(linktlist)
print(len(linktlist))

option = Options()
option.headless = True
driver = webdriver.Firefox()

profile = webdriver.FirefoxProfile()

profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", "/home/manthila/Downloads")
profile.set_preference("browser.helperApps.alwaysAsk.force", False)
profile.set_preference("browser.download.manager.alertOnEXEOpen", False)
profile.set_preference("browser.download.manager.focusWhenStarting", False)
profile.set_preference("browser.download.manager.useWindow", False)
profile.set_preference("browser.download.manager.showAlertOnComplete", False)
profile.set_preference("browser.download.manager.closeWhenDone", False)
profile.set_preference("browser.helperApps.neverAsk.openFile","pdf")
profile.set_preference("browser.helperApps.neverAsk.saveToDisk","pdf")
w = 1
for x in linktlist[2:]:
    try:
        driver.get(x)
        time.sleep(6)
        main_window = driver.current_window_handle
        time.sleep(10)

        driver.find_element_by_xpath("/html/body/div[4]/main/article[1]/div/div/div[2]/div/div/a").click()

        time.sleep(4)
        window_after = driver.window_handles[w]
        w = w + 1

        #switch on to new child window
        driver.switch_to.window(window_after)
        time.sleep(15)
        driver.find_element_by_id("download").click()
        time.sleep(15)

        print(x)

    except selenium.common.exceptions.NoSuchElementException:
        continue



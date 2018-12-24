import xlwt
import xlrd
import selenium.webdriver as webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

# creates excel files for the MSC classification numbers
wb1 = xlrd.open_workbook("msc.xls")
ws1 = wb1.sheet_by_name("Sheet 1")

wb2 = xlwt.Workbook()
ws2 = wb2.add_sheet("Sheet 1")

# creates the webdriver that points to MathSciNet
driver = webdriver.Chrome("C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe")
driver.get("https://mathscinet-ams-org.libezp.lib.lsu.edu/mathscinet/index.html")

# a redirection to a LSU login site occurs, so point the webdriver to it
# window_handles is a list of the tabs in the browser, [0] is the leftmost tab
driver.switch_to.window(driver.window_handles[0])

# waits for the LSU login site to be loaded before entering my login information
wait = WebDriverWait(driver, 10)
element = wait.until(ec.element_to_be_clickable((By.ID, 'username')))
driver.find_element_by_xpath("//input[@id='username']").send_keys("username")
driver.find_element_by_xpath("//input[@id='password']").send_keys("password" + Keys.ENTER)

# points the webdriver to the leftmost tab after logging in, now MathSciNet
driver.switch_to.window(driver.window_handles[0])

for i in range(1, 500):
    driver.find_element_by_xpath('//*[@id="publications"]/div[3]/select[1]/option[5]').click()
    driver.find_element_by_xpath('//*[@id="publications"]/div[2]').send_keys(ws1.cell(i + 500, 1).value)
    driver.find_element_by_xpath('//*[@id="publications"]/div[3]').send_keys(ws1.cell(i + 500, 2).value + Keys.ENTER)
    try:
        ws2.write(i, 1, driver.find_element_by_xpath("//a[starts-with(@href,'/mathscinet/search/mscdoc.html?code=')]").text)
        driver.find_element_by_xpath("//*[@id='logo']/img").click()
    except:
        ws2.write(i, 1, " ")
        driver.find_element_by_xpath("//*[@id='logo']/img").click()
    
wb2.save("Highly Cited Math Inclusive for Aaron 2.xls")

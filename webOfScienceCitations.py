import xlrd
import xlwt
import selenium.webdriver as webdriver
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#creates the webdriver that points to WebOfScience
driver = webdriver.Chrome("C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe")
driver.get("http://apps.webofknowledge.com.libezp.lib.lsu.edu/WOS_GeneralSearch_input.do?product=WOS&search_mode=GeneralSearch&SID=6Bgyo1mxBIsY9gH4KR5&preferencesSaved=")

#a redirection to a LSU login site occurs, so point the webdriver to it
#window_handles is a list of the tabs in the browser, [0] is the leftmost tab
driver.switch_to.window(driver.window_handles[0])

#waits for the LSU login site to be loaded before entering my login information
wait = WebDriverWait(driver, 10)
element = wait.until(EC.element_to_be_clickable((By.ID, 'username')))
driver.find_element_by_xpath("//input[@id='username']").send_keys("username")
driver.find_element_by_xpath("//input[@id='password']").send_keys("password" + Keys.ENTER)

#points the webdriver to the leftmost tab after logging in, now MathSciNet
driver.switch_to.window(driver.window_handles[0])

#opens the Excel workbook to access MSN data
wb1 = xlrd.open_workbook("msnCitations10.xls")
ws1 = wb1.sheet_by_name("Sheet 1")

#opens the Excel workbook to input WoS data
wb2 = xlwt.Workbook()
ws2 = wb2.add_sheet("Sheet 1")
ws2.write(0, 0, "Title")
ws2.write(0, 1, "Date")
ws2.write(0, 2, "WoS Citations")
ws2.write(0, 3, "WoS Classificataion")

#resets the form and adds a search row
driver.find_element_by_xpath("//*[@id=\"addSearchRow1\"]/a").click()

#loops through the MSN titles
for i in range(1, 54):

    #search for the paper by title
    driver.find_element_by_xpath("//*[@id='select2-select1-container']").click()
    driver.find_element_by_xpath("/html/body/span[34]/span/span[1]/input").send_keys("Title" + Keys.ENTER)
    driver.find_element_by_xpath("//*[@id='value(input1)']").send_keys(ws1.cell(i,0).value)

    #search for the paper by publication year
    driver.find_element_by_xpath("//*[@id='select2-select2-container']").click()
    driver.find_element_by_xpath("/html/body/span[34]/span/span[1]/input").send_keys("Year Published" + Keys.ENTER)
    driver.find_element_by_xpath("//*[@id='value(input2)']").send_keys(ws1.cell(i,2).value)

    #clicks the search button
    driver.find_element_by_xpath("//*[@id='searchCell2']/span[1]/button").click()

    #handles the exception that the paper is not found in WoS
    try:
        
        #opens the result in a new tab
        result = driver.find_element_by_class_name("smallV110")
        ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.CONTROL).click(result).key_up(Keys.CONTROL).key_up(Keys.SHIFT).perform()

        #switches the window 
        driver.switch_to.window(driver.window_handles[1])
            
        #writes to the Excel file
        ws2.write(i, 2, driver.find_element_by_class_name("large-number").text)
        ws2.write(i, 3, driver.find_element_by_xpath("//*[contains(text(), 'Web of Science Categories:')]/..").text[26:])

        #since we expect the search for "Your search found no records." to sometimes throw a NoElementException, this closes the window when it does
        driver.close()

        #points the driver back to the main WoS website
        driver.switch_to.window(driver.window_handles[0])

        #goes back to the main search page
        driver.find_element_by_xpath("/html/body/div[1]/h1/div/a/span").click()

    except:

        #writes nothing to the Excel file
        ws2.write(i, 1, " ")
        ws2.write(i, 2, " ")
        ws2.write(i, 3, " ")

        #goes back to the main search page
        driver.find_element_by_xpath("/html/body/div[1]/h1/div/a/span").click()

    #resets the form and adds a search row
    driver.find_element_by_xpath("//*[@id='addSearchRow2']/span/span[2]").click()
    driver.find_element_by_xpath("//*[@id='addSearchRow1']/a").click()
   
wb2.save("wosCitations10.xls")

    

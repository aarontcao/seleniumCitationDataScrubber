import xlwt
import selenium.webdriver as webdriver
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#creates a list and excel file for the MSC classification numbers
mscLinks = []
wb = xlwt.Workbook()
ws = wb.add_sheet("Sheet 1")
ws.write(0, 0, "Title")
ws.write(0, 1, "Journal")
ws.write(0, 2, "Year")
ws.write(0, 3, "Math Review #")
ws.write(0, 4, "MSN Citations")
ws.write(0, 5, "WoS Citations")
ws.write(0, 6, "MSN Classification")
ws.write(0, 7, "WoS Classificataion")

#creates the webdriver that points to MathSciNet
driver = webdriver.Chrome("C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe")
driver.get("https://mathscinet-ams-org.libezp.lib.lsu.edu/mathscinet/search/publications.html?arg3=&co4=AND&co5=AND&co6=AND&co7=AND&dr=all&extend=1&pg4=AUCN&pg5=TI&pg6=PC&pg7=ALLF&pg8=ET&review_format=html&s4=&s5=&s6=&s7=%22FEATURED%20REVIEW%22&s8=Journals&sort=Newest&vfpref=html&yearRangeFirst=&yearRangeSecond=&yrop=eq&r=901")

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

#changes the citations to EndNote
driver.find_element_by_xpath("//select[@name='fmt']/option[text()='Citations (EndNote)']").click()

#makes lists of all of the papers and their checkboxes
journals = driver.find_elements_by_xpath("//a[@class='mrnum' and @title='Full MathSciNet Item']")
checkboxes = driver.find_elements_by_class_name("checkbox")

#fills (rather inefficiently) the list of headline texts
for x in range(2,102):
    mscLinks.extend(driver.find_elements_by_xpath("//*[@id='content']/form/div[3]/div[2]/div/div/div[%d]/div[2]/a[last()]"%x))
    
#loops through every journal
count = 0
for x in range(57):
    
    #points the driver back to the main MathSciNet website where all of the paper links are
    driver.switch_to.window(driver.window_handles[0])

    checkboxes[x].click()
    
    #opens a paper in a new tab
    ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.CONTROL).click(journals[x]).key_up(Keys.CONTROL).key_up(Keys.SHIFT).perform()
    
    #points the driver to the website in the new tab
    driver.switch_to.window(driver.window_handles[1])
    
    #we expect searching for "FEATURED REVIEW" to throw an exception, as not every paper is featured
    try:
        if(driver.find_element_by_xpath("//span[@class='searchHighlight']").text == "FEATURED REVIEW"):

            #makes note of the number of citations from references
            citations = driver.find_element_by_xpath("//*[@id='content']/div[5]/div[1]/p[1]/a").text[17:]

            #closes the window
            driver.close()
            
            #points the driver back to the main MathSciNet website
            driver.switch_to.window(driver.window_handles[0])
            
            #since not every paper is featured, we only wish to click the checkboxes of the papers that are
            checkboxes[x].click()

            #notes the paper's MSC classification number and citations in the Excel file
            ws.write(x + 1, 4, citations)
            ws.write(x + 1, 6, mscLinks[x].text[:6])

            #makes note of how many hits there are
            count += 1
        else:
            
            #closes the window anyway if the phrase is not found
            driver.close()
    except:
        
        #since we expect the search for "FEATURED REVIEW" to sometimes throw a NoElementException, this closes the window when it does
        driver.close()

#points the driver back to the main MathSciNet website
driver.switch_to.window(driver.window_handles[0])

#generates the citations
driver.find_element_by_link_text("Retrieve Marked").click()

#points the driver back to the new citation website
driver.switch_to.window(driver.window_handles[0])

#shows all the results, will throw an exception if not enough to generate the link
try:
    driver.find_element_by_link_text("Show all results").click()
except:
    pass

#notes the indentification data in the Excel file
for x in range(1, count+1):
    for line in driver.find_element_by_xpath("//*[@id='content']/div[2]/pre[%d]"%x).text.splitlines():
        parts = line.split(" ")
        if (parts[0] == "%T"):
            title = " ".join(parts[1:])
        elif (parts[0] == "%J"):
            journal = " ".join(parts[1:])
        elif (parts[0] == "%D"):
            year = " ".join(parts[1:])
        elif (parts[0] == "%L"):
            reviewNum = " ".join(parts[1:])
    ws.write(x, 0, title)
    ws.write(x, 1, journal)
    ws.write(x, 2, year)
    ws.write(x, 3, reviewNum)
    
#saves the Excel file
wb.save("manualCheck1.xls")

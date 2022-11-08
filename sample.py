
#find case status text from web portal

#from https://preethamdpg.medium.com/running-selenium-webdriver-with-python-on-an-aws-ec2-instance-be9780c97d47
#Sample.py
#print out HMTL element from Google Home Page
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
import time
import pandas as pd
import xlrd, xlwt
import win32com.client as win32
from webdriver_manager.chrome import ChromeDriverManager




options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("enable-automation")
options.add_argument("--disable-infobars")
options.add_argument("--disable-dev-shm-usage")

service = ChromeService(executable_path=r"PATH_TO_CHROMEDRIVER")
driver = webdriver.Chrome(service=service, options=options)


#generate list of open cases to work from
excel = win32.Dispatch('excel.application')
wb = excel.workbooks.Open(r"PATH_TO_XLSX_FILE")
readData = wb.worksheets('Sheet1')
allData = readData.UsedRange
myRange = allData.Range('A1:A7')

#create list from excel file
case_list = []
def createlist():
    master_case_list = []
    for row in myRange:
        for cell in row:
            if cell.value:
                master_case_list = int(cell.value)
                case_list.append(master_case_list)
    print('Data on selected sheet :' , case_list)


               
#Set the write range to define where data will be written to
writeRange = allData.Range('B1:B7')

#create lists to save data into


date_results = []
type_results = []
room_results = []
status_results =[]


#remove html tags

def remove_html_tags(text):
    """Remove html tags from a string"""
    import re
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)




final_df = pd.DataFrame()

def search_bot(i):
    driver.get("https://web6.seattle.gov/courts/ECFPortal/default.aspx")
    driver.find_element(By.XPATH, "//li[4]/a/span/span/span").click()
    time.sleep(7)
    driver.find_element(By.XPATH, "//div[@id='ContentPlaceHolder1_CaseInfo1_CaseSearch1_pnlSubmit']/input").click()
    driver.find_element(By.XPATH, "//div[2]/div/div[3]/div/input").send_keys(str(i))
    driver.find_element(By.XPATH, "//div[2]/div/div[3]/div/input[2]").click()
    driver.find_element(By.XPATH, "//div[2]/div/ul/li[3]/a/span/span/span").click()
    time.sleep(6)
    date = driver.find_element(By.XPATH, "//div[2]/div/div/table/tbody/tr/td").get_attribute("outerHTML")


    room = driver.find_element(By.XPATH, "//div[2]/div/div/table/tbody/tr/td[3]").get_attribute("outerHTML")
    statush = driver.find_element(By.XPATH, "//div[2]/div/div/table/tbody/tr/td[4]").get_attribute("outerHTML")
    hearing_type = driver.find_element(By.XPATH, "//div[2]/div/div/table/tbody/tr/td[2]").get_attribute("outerHTML")
    date_results.append(remove_html_tags(date))
    type_results.append(remove_html_tags(hearing_type))
    room_results.append(remove_html_tags(room))
    status_results.append(remove_html_tags(statush))
    



    #element_text = driver.page_source
    #print(element_text)
    
#run functions        
createlist()   
print("Added the following hearing dates to spreadsheet: ")
for i in case_list:
    search_bot(i)

dates_df2 = pd.DataFrame()
dates_df2['Case Number'] = case_list
dates_df2['Hearing Date'] = date_results
dates_df2['Hearing Type'] = type_results
dates_df2['Room Number'] = room_results
dates_df2['Hearing Status'] = status_results
 
  
#Returns Dataframe with Case Number, Hearing Date, Room, Status and Type for each column  
print(dates_df2)

    
for row in myRange:
        for cell in row:
            if cell.value:
                for cell in row:
                    for status_result in status_results:
                        writeRange.Value = status_results
                                  
dates_df2.to_csv(r"DESIRED_OUTPUT_PATH")

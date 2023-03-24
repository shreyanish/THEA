import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

import csv

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time


#connecting google sheets
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("findLinkedin/creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("pytest").sheet1 

#initianting web driver
ser = Service("/Users/shreyanish/Dev/pythonXmentorrelations/findLinkedin/chromedriver.exe")
op = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=ser, options=op)
wait = WebDriverWait(driver, 30)

#signing into linkedin

def search_profile(name, company, i):

    search_term = "site:linkedin.com/in/ and {Name} {Company} people".format(Name = name, Company = company)
    print(search_term)

    search_bar = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input'))
    )

    search_bar.clear()
    search_bar.send_keys(search_term)
    search_bar.send_keys(Keys.ENTER) 
    time.sleep(2)
    
    profile_div = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input'))
    )
    time.sleep(5)
    profile_url = profile_div.get_attribute('href')
    print(profile_url)
    
    sheet.update_cell(i,5,profile_url)
    

with open("findLinkedin/pytest - Sheet2 (1).csv", "r") as file:
    csvreader = csv.DictReader(file)

    #signin_linkedin()

    driver.get("https://www.google.com/search?q=conquest+bits+pilani&oq=conquest+bits+pilani&aqs=chrome..69i57j0i390.4657j0j7&sourceid=chrome&ie=UTF-8")
    rowcount = 230
    
    for row in csvreader:
        Name = row['Name']
        print(Name)
        Company = row['Company']
        print(Company)
        time.sleep(5)
        try:
            search_profile(name = Name, company = Company, i=rowcount)
        except:
            continue
        
        rowcount += 1
    

    

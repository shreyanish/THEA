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

def search_profile(name, company, i = 1):

    search_term = "site:linkedin.com/in/ and {Name} {Company} and people".format(Name = name, Company = company)
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
    

positions = ['head of corporate communications', 'executive assistant', 'head of brand communications','head of pr', 'head of marketing']


def search_profile_events(company):
    
    for j in range(len(positions)-1):

        search_term = "site:linkedin.com/in/ and {Company} {Position} and people".format(Position = positions[j], Company = company)
        print(search_term)
        search_bar = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input'))
        )

        search_bar.clear()
        search_bar.send_keys(search_term)   
        search_bar.send_keys(Keys.ENTER) 
        time.sleep(2)

        profile_div = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="rso"]/div[1]/div/div/div[1]/div/a'))
        )
        time.sleep(5)
        profile_url = profile_div.get_attribute('href')
        print(profile_url)

        profile_info = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="rso"]/div[1]/div/div/div[1]/div/a/h3'))
        )
        print(profile_info)

        insertRow = [profile_info.text, company, profile_url]
        sheet.insert_row(insertRow, 115) 
        

positions_mnc = ['brand strategy', 'marketing head', 'strategic partnerships & alliance', ]

def search_profile_mncs(company):
    
    for j in range(len(positions_mnc)-1):

        search_term = "site:linkedin.com/in/ and {Company} present {Position} and people india".format(Position = positions_mnc[j], Company = company)
        print(search_term)
        search_bar = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input'))
        )

        search_bar.clear()
        search_bar.send_keys(search_term)   
        search_bar.send_keys(Keys.ENTER) 
        time.sleep(2)

        profile_div = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="rso"]/div[1]/div/div/div[1]/div/a'))
        )
        time.sleep(5)
        profile_url = profile_div.get_attribute('href')
        print(profile_url)

        profile_info = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="rso"]/div[1]/div/div/div[1]/div/a/h3'))
        )
        print(profile_info)

        insertRow = [profile_info.text, company, profile_url]
        print('InsertRow: ',insertRow)
        sheet.insert_row(insertRow, 1) 


with open("findLinkedin/campus.csv", "r") as file:
    csvreader = csv.DictReader(file)

    #signin_linkedin()

    driver.get("https://www.google.com/search?q=conquest+bits+pilani&oq=conquest+bits+pilani&aqs=chrome..69i57j0i390.4657j0j7&sourceid=chrome&ie=UTF-8")
    
    
    for row in csvreader:
        Company = row['Company']
        print(Company)
        time.sleep(5)
        try:
            search_profile_mncs(company=Company) 
        except:
            continue

        

        
        
    

    

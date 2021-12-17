import pandas as pd
import pandas.io.formats.excel
pandas.io.formats.excel.ExcelFormatter.header_style = None
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook, load_workbook
import os
import os.path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import glob
from selenium.webdriver.common.by import By
os.environ['PATH'] += r";C:\SeleniumDriver"

### Initial Scraping

driver = webdriver.Chrome()
driver.get("https://gatfl.gatech.edu/sri/users/login")
driver.maximize_window()


username = driver.find_element(By.ID, "UserUsername")
username.clear()
username.send_keys("nsandell6")


password = driver.find_element(By.ID, "UserPassword")
password.clear()
password.send_keys("nsandell123!")


button = driver.find_element(By.XPATH, "//input[@type='submit']")
button.click()

requests_link = driver.find_element(By.LINK_TEXT, "GA-DOE Requests")
requests_link.click()

export_requests = driver.find_element(By.LINK_TEXT, 'Export all')
export_requests.click()

driver.get('https://gatfl.gatech.edu/sri/users')

export_requests = driver.find_element(By.LINK_TEXT, 'Export GA-DoE')
export_requests.click()


### Collect CSV from downloads
date = str(input("What is today's date in (MMDDYYYY)"))
folder_path = r'C:\Users\sande\Downloads'
file_type = '\*csv'
files = glob.glob(folder_path + file_type)

files = sorted(files, key = os.path.getctime, reverse=True)
base_name = r"C:\Users\sande\Downloads\GA DoE "
new_requests_name = base_name + "Requests - " + date + ".csv"
new_users_name = base_name + "Users - " + date + ".csv"
os.rename(files[0], new_users_name)
os.rename(files[1], new_requests_name)

### Insert into Excel

df_requests = pd.read_csv(new_requests_name)
df_requests.to_excel('requests_output.xlsx', index=False)

df_users = pd.read_csv(new_users_name)
df_users.to_excel('users_output.xlsx', index=False)


wb = load_workbook(filename='insert.xlsm', read_only=False, keep_vba=True)

ws_requests = wb.create_sheet(index=1)
ws_requests.title = "GA DoE Requests - " + date
ws_users = wb.create_sheet(index=4)
ws_users.title = "GA DoE Users - " + date

for row in dataframe_to_rows(df_requests, index=True, header=True):
    ws_requests.append(row)

ws_requests.delete_rows(2)
ws_requests.delete_cols(1)

for row in dataframe_to_rows(df_users, index=True, header=True):
    ws_users.append(row)
ws_users.delete_rows(2)
ws_users.delete_cols(1)


wb.remove(wb[wb.sheetnames[3]])
wb.remove(wb[wb.sheetnames[5]])

new_file_name = os.getcwd() + "\\" + "Combined Data " + date + ".xlsm"
wb.save(filename=new_file_name)

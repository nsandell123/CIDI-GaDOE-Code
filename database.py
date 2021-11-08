from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
os.environ['PATH'] += r";C:\SeleniumDriver"

driver = webdriver.Chrome()
driver.get("https://gatfl.gatech.edu/sri/users/login")
driver.maximize_window()


username = driver.find_element_by_id("UserUsername")
username.clear()
username.send_keys("TEST")

password = driver.find_element_by_id("UserPassword")
password.clear()
password.send_keys("TEST")

button = driver.find_element_by_xpath("//input[@type='submit']")
button.click()

requests_link = driver.find_element_by_link_text('GA-DOE Requests')
requests_link.click()

export_requests = driver.find_element_by_link_text('Export all')
export_requests.click()

driver.get('https://gatfl.gatech.edu/sri/users')

export_requests = driver.find_element_by_link_text('Export GA-DoE')
export_requests.click()
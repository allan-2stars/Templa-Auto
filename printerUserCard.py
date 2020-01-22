# import package

from selenium import webdriver
from selenium.common.exceptions import *
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

import time

driver = webdriver.Chrome()
driver.get("http://192.168.10.19/account_userlist.html")
# 192.168.10.12
# 192.168.10.19
# 192.168.10.20
# 192.168.10.21

## if the login page is with admin login page, then use below
## try to find the special login page button
def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

## due to same xpath for different login setting page, so 
## check and use are separate xpath
xpath_to_check = '//*[@id="mainContentsScroll"]/form/div[1]/div/input[3]'
xpath_to_use = '//*[@id="mainContentsScroll"]/form/div[1]/div/input[2]'

print(str(check_exists_by_xpath(xpath_to_check)))
if check_exists_by_xpath(xpath_to_check):  
    
    ## click the login admin button
    driver.find_element_by_xpath(xpath_to_use).click()
    
    ## find the password textbox
    password_input = driver.find_element_by_xpath('//*[@id="element10006"]')
else:
    ## find the password textbox directly
    password_input = driver.find_element_by_xpath('//*[@id="element10003"]')

password_input.send_keys('admin')
login_button = driver.find_element_by_xpath('//*[@id="mainContentsScroll"]/form/div[3]/div/input[1]')
login_button.click()

## after login

## get the User List Page, and show all 100 per page
userlist_menu_button = driver.find_element_by_xpath('//*[@id="currentmenu"]/a')
userlist_menu_button.click()
user_index_button = driver.find_element_by_xpath('//*[@id="main"]/form/table[3]/tbody/tr[2]/td[2]/select')
user_index_button.click()

## show 100 users per page
display_items_list = driver.find_element_by_xpath('//*[@id="main"]/form/table[3]/tbody/tr[3]/td[2]/select')
display_items_list.send_keys("100")


#######################################
#
# Setup Excel Sheet to read user names
#
#######################################
sheetLoader = '192.168.10.19' 
df = pd.read_excel('printer-setup.xlsx', sheetname=sheetLoader)
print("starting...")

## literate all the record from Excel sheet
## read code and name from the list and apply them into the printer website
for i in df.index:
    user_code = df['CODE']
    user_name = df['NAME']
    status = df['STATUS']

    if status[i] == "Done" or status[i] == "Skip":
        print(str(user_code[i]) + ', done for ' + user_name[i])
        continue

    if status[i] == "Stop":
        print("Stop here")
        break


    ## Search Users one by one and change the settings
    username_link = driver.find_element_by_link_text(user_name[i].strip())
    username_link.click()
    ## on the newly opened page, click on submit to setup the Card ID
    driver.find_element_by_xpath('//*[@id="main"]/form/table[1]/tbody/tr[5]/td/input').click()
    
    ############################################
    ##
    ## Change the Card ID
    ## Need to detect "Delete" button, disabled or enabled
    ##
    ############################################
    ## check if the "Delete" button is enabled
    
    ## pause to give time to open the new page
    
    ## parse all code to text and 
    ## adding leading 0 at front
    user_code_text = '0'+ str(user_code[i])
    time.sleep(2)
    
    register_delete_button = driver.find_element_by_xpath('//*[@id="main"]/form/table[2]/tbody/tr/td/input')
    print('detecting the delete button ...')
    
    if register_delete_button.is_enabled():
        register_delete_button.click()
        print('deleted the old record ...')
        
    ## once check the textbox is ediable by checking the delete button status,
    ## then type in the Card ID code
    register_card_id_textbox = driver.find_element_by_xpath('//*[@id="element14"]')
    register_card_id_textbox.send_keys(str(user_code_text))
    print('input the code number ...')
    
    # Click OK button
    register_ok_button = driver.find_element_by_xpath('//*[@id="mainContentsScroll"]/div[4]/div/input[1]')
    register_ok_button.click()

    ## click the "Submit" button to finalise the setup
    submit_button = driver.find_element_by_xpath('//*[@id="mainContentsScroll"]/div[2]/div/input[1]')
    submit_button.click()
    print('done for: ' + user_name[i])
    time.sleep(2)

#driver.close()
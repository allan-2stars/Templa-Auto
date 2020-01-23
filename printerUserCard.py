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

def check_exists_by_link_text(link_name):
    try:
        driver.find_element_by_link_text(link_name)
    except NoSuchElementException:
        return False
    return True

## due to same xpath for different login setting page, so 
## check and use are separate xpath
xpath_to_check = '//*[@id="mainContentsScroll"]/form/div[1]/div/input[3]'
xpath_to_use = '//*[@id="mainContentsScroll"]/form/div[1]/div/input[2]'

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
sheetLoader = 'id_list' 
df = pd.read_excel('printer-setup.xlsx', sheet_name=sheetLoader)
print("starting...")

## literate all the record from Excel sheet
## read code and name from the list and apply them into the printer website
for i in df.index:
    user_code = df['CODE']
    user_name = df['NAME']
    user_email = df['EMAIL']
    user_setup = df['SETUP']
    status = df['STATUS']
    
    ## convert numbers to string text
    ## parse all code to text and 
    ## adding leading 0 at front
    ## uncommon below to use code with leading 0
    # user_code_text = '0'+ str(int(user_code[i]))
    ## uncommon below to use code without leading 0
    user_code_text = str(int(user_code[i]))  

    if status[i] == "Done" or status[i] == "Skip":
        print(user_code_text + ', DONE for ' + user_name[i])
        continue
        
    if status[i] == "Skip":
        print(user_code_text + ', SKIPPED for ' + user_name[i])
        continue

    if status[i] == "Stop":
        print("STOPPED as Files told!!")
        break


    ## Search Users one by one and change the settings
    

    ################################################
    ##
    ## Check if the User Name link text exists
    ##
    ################################################
    
    ## as long as the link text cannot be found, then click next page
    ## if till the last page, still unfound, then set the link_name_found as False.
    link_text_found = True  ## a flag for check if found the user link
    next_page_button = driver.find_element_by_xpath('//*[@id="main"]/form/p[1]/input[2]')
    
    ## Search the user name link until found, 
    while not check_exists_by_link_text(user_name[i].strip()):
        ## if not found, see if this is the last page?
        next_page_button = driver.find_element_by_xpath('//*[@id="main"]/form/p[1]/input[2]')
        
        ## if this is the last page, the stop the loop
        ## set the flag to False
        if not next_page_button.is_enabled():
            link_text_found = False
            break
            
        ## if not the last page, go to next page
        else:
            next_page_button.click()


    print('--- finished checking user name ---')
    ## if user link exists, click on the link
    if link_text_found == True:
        username_link = driver.find_element_by_link_text(user_name[i].strip())
        username_link.click()
    ## if the user link not found, continue to check next user
    else:
        print('--------------- Alert -------------')
        print('User '  + user_name[i] + ' Cannot be Found !!!')
        print(' ')
        print(' --------- Very Important !!------------------')
        print(' ---------------------------------------------')
        continue
        
    ################################################
    ##
    ## Click on 'Submit' button to setup the Card ID
    ##
    ## #############################################
    submit_button_xpath = '//*[@id="main"]/form/table[1]/tbody/tr[6]/td/input'
    if check_exists_by_xpath(submit_button_xpath):
        driver.find_element_by_xpath(submit_button_xpath).click()
    else:
        driver.find_element_by_xpath('//*[@id="main"]/form/table[1]/tbody/tr[5]/td/input').click()
    
    ################################################
    ##
    ## Update the Card ID
    ## Need to detect "Delete" button, disabled or enabled
    ##
    ################################################
    ## check if the "Delete" button is enabled
    
    ## pause to give time to open the new page
       
    time.sleep(2)
    
    register_delete_button = driver.find_element_by_xpath('//*[@id="main"]/form/table[2]/tbody/tr/td/input')
    print('detecting the delete button ...')
    
    if register_delete_button.is_enabled():
        register_delete_button.click()
        print('deleted the old record ...')
        
    ## once check the textbox is ediable by checking the delete button status,
    ## then type in the Card ID code
    register_card_id_textbox = driver.find_element_by_xpath('//*[@id="element14"]')
    register_card_id_textbox.send_keys(user_code_text)
    print('input the code number ...')
    
    # Click OK button
    register_ok_button = driver.find_element_by_xpath('//*[@id="mainContentsScroll"]/div[4]/div/input[1]')
    register_ok_button.click()
    
    ################################################
    ##
    ## Update the User Email
    ## Need to detect if email exist or not
    ##
    ################################################
    
    
    email_textbox_xpath = '//*[@id="element8"]'
    email_textbox = driver.find_element_by_xpath(email_textbox_xpath)
    if check_exists_by_xpath(email_textbox_xpath):
        ## if email not exists in system, then enter the email
        if email_textbox.get_attribute('value') == '':
            email_textbox.send_keys(user_email[i])
            print('email updated ...')
        else: ## if email already exists, skip
            print('email already exist, no changes for email ...')
    else:
        print('email textbox not found, SKIP Email entry only !!')
    ## if some else email xpath exist later, uncommon below,
    ## and change the xpath to the real one
    ## else:
    ##     driver.find_element_by_xpath('//*[@id="element8"]"]').send_keys(user_email[i])
    ## if the email textbox xpath changed to another one, use below
   

        
    ################################################  
    ##
    ## Once finish all the works, 
    ## then submit the form to save the changes
    ##
    ################################################
    ## click the "Submit" button to finalise the setup
    submit_button = driver.find_element_by_xpath('//*[@id="mainContentsScroll"]/div[2]/div/input[1]')
    submit_button.click()
    print('done for: ' + user_name[i])
    time.sleep(2)
    
print('-------------- Done all ------------')
    

#driver.close()
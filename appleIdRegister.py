# import package

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time

#######################################################
### all web elements/handles info recoreded as below
xpath_first_name = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[2]/div/div[1]/div/div/full-name/div[1]/div/div/first-name-input/div/idms-textbox/idms-error-wrapper/div/div/input'
xpath_last_name = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[2]/div/div[1]/div/div/full-name/div[2]/div/div/last-name-input/div/idms-textbox/idms-error-wrapper/div/div/input'
xpath_countries = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[2]/div/div[2]/div/idms-dropdown/div/idms-error-wrapper/div/div/select'

xpath_dob = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[2]/div/div[3]/div/wc-birthday/div/div/div/div/masked-date/idms-error-wrapper/div/div/input'
xpath_email = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[3]/div/div[1]/div/account-name/div/div/email-input/div/idms-textbox/idms-error-wrapper/div/div/input'
xpath_password = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[3]/div/div[2]/div/new-password/div/div/password-input/div/input'
xpath_password_confirm = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[3]/div/div[2]/div/confirm-password/div/div/confirm-password-input/div/idms-textbox/idms-error-wrapper/div/div/input'
xpath_region_code_selection = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[4]/div/div/div/div/phone-number/div/div/div[1]/div/phone-input/div/idms-dropdown/div/idms-error-wrapper/div/div/select'
xpath_phone_number = '/html/body/div[2]/aid-web/div[2]/div[2]/div/create-app/aid-create/idms-flow/div/div/div/idms-step/div/div/div/div[2]/div/div/div[4]/div/div/div/div/phone-number/div/div/div[1]/div/phone-input/div/idms-textbox/idms-error-wrapper/div/div/input'
announcements_tick_id = 'news'
apps_more_tick_id = 'itunes'
apple_news_tick_id = 'appleNews'
#######################################################


## apple change id xpath after refreshing the webpage.
## the only solution here is use tab key

####################################################################
### Get data from Excel Sheet, loop over until all entries are done
####################################################################
site_reallocate_sheet = 'users' 
df = pd.read_excel('apple_register.xlsx', sheet_name=site_reallocate_sheet)
print("starting...")
driver = webdriver.Chrome()
browser_tab_index_offset = 0 ## offset the index number once meet a Done or Skip

for i in df.index:

    ### get the full name of the user
    first_name = df['First Name']
    last_name = df['Last Name']
    full_name_string = str(first_name[i]) + ' ' + str(last_name[i])

    ### if the status is Done, then skip to next entry
    status = df['Status']
    if status[i] == "Done":
        browser_tab_index_offset +=1
        print(full_name_string  + " is Done")
        continue
    if status[i] == "Skip":
        browser_tab_index_offset += 1
        print(full_name_string + " is Skipped")
        continue
    if status[i] == "Stop":
        print("Stop here")
        break
    
    ###############################################################
    ### Get ready for data from Excel
    ###############################################################

    country = df['Country']
    dob = df['DOB']
    email = df['Email']
    password = df['Password']
    password_2 = df['Password_2']
    phone_code = df['Phone Code']
    phone_number = df['Phone Number']
    URL = df['URL']


 
    try:
        dob_string = format(int(float(dob[i])),'08d')
        phone_code_string = str(phone_code[i])
        phone_number_string = format(int(float(phone_number[i])),'010d')
    except:
        print('Program stopped, check if all user registered propurly ...')
        break

    #################################################################

    #######
    ## except for the first run, you need open a new tab for multiple user register
    if i > 0:
        driver.execute_script("window.open();")
        driver.switch_to.window(driver.window_handles[i - browser_tab_index_offset])
    ## open the URL
    driver.get(URL[i])

    ##############################################
    ### get the elements from the browser
    ##############################################
    first_name_element = driver.find_element_by_xpath(xpath_first_name)
    last_name_element = driver.find_element_by_xpath(xpath_last_name)
    countries_element = driver.find_element_by_xpath(xpath_countries)
    dob_element = driver.find_element_by_xpath(xpath_dob)
    email_element = driver.find_element_by_xpath(xpath_email)
    password_element = driver.find_element_by_xpath(xpath_password)
    password_2_element = driver.find_element_by_xpath(xpath_password_confirm)
    region_code_selection_element = Select(driver.find_element_by_xpath(xpath_region_code_selection))
    phone_number_element = driver.find_element_by_xpath(xpath_phone_number)

    announcements_tick_element = driver.find_element_by_id(announcements_tick_id)
    apps_more_tick_element = driver.find_element_by_id(apps_more_tick_id)
    apple_news_tick_element = driver.find_element_by_id(apple_news_tick_id)
    ##############################################


    ######################################################
    ### Operate the element fill text or select options
    ######################################################
    first_name_element.send_keys(first_name[i])
    last_name_element.send_keys(last_name[i])
    countries_element.send_keys(country[i])
    dob_element.send_keys(dob_string)
    email_element.send_keys(email[i])
    password_element.send_keys(password[i])
    password_2_element.send_keys(password_2[i])
    region_code_selection_element.select_by_index(11)
    phone_number_element.send_keys(phone_number_string)

    if announcements_tick_element.get_attribute('checked'):
        announcements_tick_element.send_keys(Keys.SPACE)
    if apps_more_tick_element.get_attribute('checked'):
        apps_more_tick_element.send_keys(Keys.SPACE)
    ## some reason cannot detect below checkbox, do it manully for now.
    # if apple_news_tick_element.get_attribute('checked'):
    #     apple_news_tick_element.send_keys(Keys.SPACE)
    #######################################################




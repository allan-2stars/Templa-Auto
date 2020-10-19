import pyautogui
import pywinauto
import pandas as pd
import time
from functions.functions_utils import tm_init


#############################
##
## QA Recipients Reallocation
##
#############################

def QA_Recipients():
    if tm_init() is None:
        print("Can't find Templa on your computer")
    else:
        templa = tm_init()[0]
        app = tm_init()[1]
        ## start 
        print("Starting...")
        mainContractsTab = templa.child_window(title='Contracts', control_type='TabItem')
        mainContractsTab.click_input()
        mainContractsWindow = templa.child_window(title='Contracts', control_type='Window')

        ########################
        #
        # Setup Excel Sheet
        #
        ########################
        sheetLoader = 'QA-Recipient' 
        df = pd.read_excel('test.xlsx', sheet_name=sheetLoader)

        for i in df.index:
            siteCode = df['CODE']
            siteName = df['SITE']
            recipient1 = df['RECIPIENT1']
            recipient2 = df['RECIPIENT2']
            recipient3 = df['RECIPIENT3']
            recipient4 = df['RECIPIENT4']
            recipient5 = df['RECIPIENT5']
            check1 = df['CHECK1']
            check2 = df['CHECK2']
            check3 = df['CHECK3']
            check4 = df['CHECK4']
            check5 = df['CHECK5']
            email1 = df['EMAIL1']
            email2 = df['EMAIL2']
            email3 = df['EMAIL3']
            email4 = df['EMAIL4']
            email5 = df['EMAIL5']
            jobTitle1 = df['JOB TITLE1']
            jobTitle2 = df['JOB TITLE2']
            jobTitle3 = df['JOB TITLE3']
            jobTitle4 = df['JOB TITLE4']
            jobTitle5 = df['JOB TITLE5']
            isFailRecipient = df['FAIL RECIPIENT']
            status = df['STATUS']
            #print("Site Name:" + siteName[i])
            #print("CSM: " + csm[i])
            #print("iPad: " + ipad[i])
            if status[i] == "Done" or status[i] == "Skip":
                print(str(siteCode[i]) + " is Done")
                continue

            if status[i] == "Stop":
                print("Stop here")
                break

            # click on the Code Edit Box
            mainContractsWindow.window(title='Site', control_type='ComboBox').click_input()
            pyautogui.typewrite(str(siteCode[i]))
            pyautogui.moveRel(0, 25) 
            pyautogui.doubleClick() # open the site by double click
            print("starting...")

            # # open analysis details dialouge window
            contractDetailWindow = app.window(title_re='Contract - *')
            contractDetailWindow.wait('exists', timeout=25)
            contractDetailWindow.window(title='QA', control_type='TabItem').click_input()

            # Check if there is QA Items
            # if no qa at all, then no need to change qa recipient

            if contractDetailWindow.window(title='30').exists():    
                print("This QA is Monthly currently." + str(siteCode[i]))
                
            elif contractDetailWindow.window(title='90').exists():     
                print("This QA is Quaterly currently." + str(siteCode[i]))

            elif contractDetailWindow.window(title='7').exists():     
                print("This QA is Weekly currently." + str(siteCode[i]))

            elif contractDetailWindow.window(title='14').exists():     
                print("This QA is Forenightly currently." + str(siteCode[i]))

            elif contractDetailWindow.window(title='365').exists():     
                print("This QA is Yearly currently." + str(siteCode[i]))

            ## if not match above, there must exist an error
            else:
                print("----------------------------------------------")       
                print("This QA is UNKNOWN frequency or QA NOT Exist." + str(siteCode[i]))
                print("Check this please ..." + str(siteCode[i]))
                print("----------------------------------------------")

                ## if you want add the recipient without checking the qa exists or not
                ## comment below 2 lines
                contractDetailWindow.Close.click_input()
                continue

            ## QA failure recipients button click.
            if str(isFailRecipient[i]).lower() == 'yes':
                contractDetailWindow['QA failure recipients'].click_input()
            ## QA recipients button click.
            else:
                contractDetailWindow['QA form recipients'].click_input()

            contractRecipientsWindow = contractDetailWindow.child_window(title_re='Contract Recipients - *')
            contractRecipientsWindow.wait('exists', timeout=135)
            
            #contractRecipientsWindow.wait('exists', timeout=15)
            print("check recipient")
            recipientExitEither = False
            isChecked = False
            recipientsAddingList = [recipient1[i]]
            
            checkStateList = [check1[i]]
            emailList = [email1[i]]
            jobTitleList = [jobTitle1[i]]

            # you need add more recipient if recipient2 in excel not empty

            if str(recipient2[i]) != "nan":
                
                recipientsAddingList.append(recipient2[i])

                checkStateList.append(check2[i])
                emailList.append(email2[i])
                jobTitleList.append(jobTitle2[i])
                print('recipient 2 added: ' + str(recipient2[i]))
            
            if str(recipient3[i]) != "nan":

                recipientsAddingList.append(recipient3[i])

                checkStateList.append(check3[i])
                emailList.append(email3[i])
                jobTitleList.append(jobTitle3[i])
                print('recipient 3 added: ' + str(recipient3[i]))

            if str(recipient4[i]) != "nan":

                recipientsAddingList.append(recipient4[i])

                checkStateList.append(check4[i])
                emailList.append(email4[i])
                jobTitleList.append(jobTitle4[i])
                print('recipient 4 added: ' + str(recipient4[i]))

            if str(recipient5[i]) != "nan":

                recipientsAddingList.append(recipient5[i])

                checkStateList.append(check5[i])
                emailList.append(email5[i])
                jobTitleList.append(jobTitle5[i])
                print('recipient 5 added: ' + str(recipient5[i]))


            arrayCount = len(recipientsAddingList)
            print ('now is row: ' + str(i+2) + ' in Excel Sheet')
            ## loop over the recipient names
            for k in range(arrayCount):
                print ("recipient's name array total: " + str(arrayCount)+ ', now is on ' + str(k+1))

                ## Need to check 2 ways around
                ## firstname lastname
                ## lastname firstname
                ## 


                nameArray = str(recipientsAddingList[k]).split(" ")
                for name in nameArray:
                    print('the name in array is: ' + name)
                if len(nameArray) == 1: ## name only contains one word
                    recipient_name = nameArray[0]
                    print('Name must be 2x words, and currently is ', recipient_name)

                if len(nameArray) == 2: ## name contains two words
                    nameFirstLast = nameArray[0] + " " + nameArray[1]
                    nameLastFirst = nameArray[1] + " " + nameArray[0]
                    recipientEntryFL = contractRecipientsWindow.window(title=nameFirstLast)
                    recipientEntryLF = contractRecipientsWindow.window(title=nameLastFirst)

                    print("check name exist or not: " + recipientsAddingList[k])
                    recipientExitEither = recipientEntryFL.exists() or recipientEntryLF.exists()
                    recipientExitBoth = recipientEntryFL.exists() and recipientEntryLF.exists()

                    checkStateOnExcel = str(int(checkStateList[k]))

                    ## Recipient Exist in the List First Name Last Name
                    
                    
                    if recipientExitBoth:
                        if checkStateOnExcel == "0":
                            print("duplicate name exist: " + nameFirstLast)
                            recipientEntryCheckboxLF = recipientEntryLF.child_window(title="Receive documents?", control_type="CheckBox")
                            recipientEntryCheckboxFL = recipientEntryFL.child_window(title="Receive documents?", control_type="CheckBox")
                            isCheckedLF = recipientEntryCheckboxLF.get_toggle_state()
                            isCheckedFL = recipientEntryCheckboxFL.get_toggle_state()
                            
                            if str(isCheckedFL) != checkStateOnExcel:
                                recipientEntryCheckboxLF.toggle()
                                recipientEntryCheckboxFL.toggle()
                                print("check state CHANGED to: " + checkStateOnExcel)
                            else:
                                print("check state SAME as before: " + checkStateOnExcel)


                        if checkStateOnExcel == "1":
                            print("duplicate name exist: " + nameFirstLast)
                            recipientEntryCheckboxLF = recipientEntryLF.child_window(title="Receive documents?", control_type="CheckBox")
                            recipientEntryCheckboxFL = recipientEntryFL.child_window(title="Receive documents?", control_type="CheckBox")
                            isCheckedLF = recipientEntryCheckboxLF.get_toggle_state()
                            isCheckedFL = recipientEntryCheckboxFL.get_toggle_state()
                            
                            # If currently not checked, let Firstname Lastname check
                            if str(isCheckedFL) != checkStateOnExcel:
                                recipientEntryCheckboxFL.toggle()
                                print("First Last, state CHANGED to: " + checkStateOnExcel)

                            # If currently checked, let Firstname Lastname remained
                            if str(isCheckedFL) == checkStateOnExcel:
                                print("First Last, state Same as: " + checkStateOnExcel)

                            # If currently checked, let Lastname Firstname off check
                            if str(isCheckedLF) == checkStateOnExcel:
                                recipientEntryCheckboxFL.toggle()
                                print("duplicate name turned off")

                            # If currently not checked, let Lastname Firstname keep un-checked  
                            if str(isCheckedLF) != checkStateOnExcel:
                                print("duplicated name no need change")

                    elif recipientEntryFL.exists():
                        print("exist: " + nameFirstLast)
                        recipientEntryCheckboxFL = recipientEntryFL.child_window(title="Receive documents?", control_type="CheckBox")
                        isCheckedFL = recipientEntryCheckboxFL.get_toggle_state()
                        
                        print('state in system now: ' + str(isCheckedFL))
                        print("  ")
                        print('recipient state should be: ' + checkStateOnExcel)
                        if str(isCheckedFL) != checkStateOnExcel:
                            recipientEntryCheckboxFL.toggle()
                            print("check state CHANGED to: " + checkStateOnExcel)
                        else:
                            print("check state SAME as before: " + checkStateOnExcel)

                    ## Recipient Exist in the List, Last Name First Name
                    elif recipientEntryLF.exists():
                        print("exist: " + nameLastFirst)
                        recipientEntryCheckboxLF = recipientEntryLF.child_window(title="Receive documents?", control_type="CheckBox")
                        isCheckedLF = recipientEntryCheckboxLF.get_toggle_state()
                        
                        print('state in system now: ' + str(isCheckedLF))
                        print("  ")
                        print('recipient state should be: ' + checkStateOnExcel)
                        if str(isCheckedLF) != checkStateOnExcel:
                            recipientEntryCheckboxLF.toggle()
                            print("check state CHANGED to: " + checkStateOnExcel)
                        else:
                            print("check state SAME as before: " + checkStateOnExcel)
                    
                    ## Recipient NOT exist, we need to add a new item
                    ## but also, need to know if add a CSM or Client
                    elif not recipientExitEither and checkStateOnExcel == "0":
                        print('recipient not exist, and no need add')
                        
                            
                    ## If need to add a CSM User
                    elif not recipientExitEither and checkStateOnExcel == "1" and jobTitleList[k] == "CSM":
                        print('recipient not exist, and need add')
                        # click on the add contact button
                        contractRecipientsWindow.child_window(title="Add user", control_type="Button").click_input()
                        # find the email
                        print("adding new CSM user to list...")
                        usersSelectWindow = app.window(title='Users')
                        usersSelectWindow.wait('exists', timeout=15)
                        usersSelectWindow.window(title='Email', control_type='ComboBox').click_input()
                        pyautogui.typewrite(str(emailList[k]))
                        pyautogui.moveRel(0, 25)
                        pyautogui.click()
                        usersSelectWindow.Select.click_input()
                        # if need more to add, continue above
                        usersSelectWindow.Close.click_input()
                        

                    ## If need to add a Client
                    elif not recipientExitEither and checkStateOnExcel == "1" and jobTitleList[k] == "Client":
                        # click on the add contact button
                        contractRecipientsWindow.child_window(title="Add contact", control_type="Button").click_input()
                        # find the email
                        print("adding new Contact to list...")
                        contactsSelectWindow = app.window(title='Contacts Select')
                        contactsSelectWindow.wait('exists', timeout=95)
                        contactsSelectWindow.window(title='Email', control_type='ComboBox').click_input()
                        pyautogui.typewrite(emailList[k])
                        pyautogui.moveRel(0, 25)
                        pyautogui.click()
                        contactsSelectWindow.Select.click_input()
                        # if need more to add, continue above
                        contactsSelectWindow.Close.click_input()
                        
                    else:
                        print('recipient exist either:', recipientExitEither)
                        print('recipient exist both:', recipientExitBoth)
                        print('need to add:', checkStateOnExcel)

                        print("something wrong, no conditions is matched")

            # Save
            contractRecipientsWindow.Save.click_input()
            time.sleep(2.5)
            contractDetailWindow.Close.click_input()
            time.sleep(2.5)
            print (str(siteCode[i]) + ": Done now")
            print ("######################################")
            print (" ")


        print ("All Done")
        print (" ")


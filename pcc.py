import time
import tkinter
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
import csv
from datetime import datetime
import time as t
import pandas as pd
from selenium.webdriver.support.ui import Select
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox
from tkinter import filedialog as fd
from threading import Thread
import traceback
import re
import requests
import socket
import json

# Global Variables
driver = None
patient_data_sheet = None
thread_stopped = False
get_e_timeout = 20
current_users = None
user_name = None
password = None
selected_sheet = ""
name_designation="Juan Garcia MD"
md_do_pa_np="PA"
clinic_or_practice="Trinity Health Services"
vis_date="12/3/2021"
temp_c=""
hr=""
bp=""
rr=""
ht=""
wt=""
no_allergies="Yes"
food="No"
medication="No"
corrected_left_eye=""
corrected_both_eyes=""
uncorrected_right_eye="20/20"
uncorrected_left_eye="20/20"
uncorrected_both_eyes="20/20"
medical_history=""
travel_history="0"
past_medical_history="Denies"
family_history="Denies"
lmp="N/A"
previous_regnancy="N/A"
no_abnormal_findings="Yes"
other_1=""
other_2=""
general_appearance="normal"
heent="normal"
neck="normal"
heart="normal"
lungs="normal"
abdomen="normal"
gu_gyn=""
describe="Deffered"
extremeties="normal"
back_spine="normal"
neurologic="normal"
skin="normal"
describe_concerns=""
mental_health="0"
h15="n"
other_medical="Contact with and (suspected) exposure to COVID-19."
value_initial_medical_exam = '549809'
value_quarantine_isolation = '547095'
value_standing_orders_12_over = '548012'


# Strings
specify_travel = "The minor is medically cleared to travel only if all covid quarantine clearance criteria have been met and no other concerns requiring medical follow up and/or specialty follow-up have been identified in subsequent visits."

def update_status(msg):
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print(current_time + " - " + msg)

def open_chrome():
    global driver
    # tkinter.messagebox.showinfo("Information", "Please log into your account using next opening driver. Then click on 'Start' button to start the automation.")
    update_status("Opening Chrome..")
    options = Options()
    options.add_argument("start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    # options.add_argument("user-data-dir=" + os.getcwd() + "/ChromeProfile")
    # driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

    # Download from https://chromedriver.chromium.org/downloads
    driver = webdriver.Chrome("chromedriver.exe", options=options)
    update_status("Navigating to CRM..")
    driver.get("https://login.pointclickcare.com/home/userLogin.xhtml")

def wait_button(d, el, t):
    try:
        el = WebDriverWait(d, get_e_timeout).until(EC.element_to_be_clickable((t, el)))
        return False
    except Exception as e:
        print(e)
        return True
        pass

def wait_window(d):
    try:
        WebDriverWait(d, get_e_timeout).until(EC.new_window_is_opened)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_text(d, el, data):
    try:
        d.find_element(By.ID, el).clear()
        d.find_element(By.ID, el).send_keys(data)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_text_name(d, el, data):
    try:
        d.find_element(By.NAME, el).clear()
        d.find_element(By.NAME, el).send_keys(data)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_click_pos(d, el, pos):
    try:
        d.find_elements(By.ID, el)[pos].click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_click_pos_by_class(d, el, pos):
    try:
        d.find_elements(By.CLASS_NAME, el)[pos].click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_click(d, el):
    try:
        d.find_element(By.ID, el).click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_click_name(d, el):
    try:
        d.find_elements(By.NAME, el).click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_enter(d, el):
    try:
        d.find_element(By.ID, el).send_keys(Keys.ENTER)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def click_link(d, el):
    try:
        d.find_element(By.XPATH, '//a[contains(text(),"' + el + '")]').click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def click_link_href(d, el):
    try:
        d.find_element(By.XPATH, '//a[@href="' + el + '"]').click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def click_button_value(d, el):
    try:
        d.find_element(By.XPATH, '//input[@value="' + el + '"]').click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def get_string_date(el):
    res = ""
    try:
        res = el.strftime("%m/%d/%Y")
    except:
        res = el
        pass
    return res

def select_window(d, pos):
    try:
        d.switch_to_window(d.window_handles[pos])
        return False
    except Exception as e:
        print(e)
        return True
        pass

def select_menu_name(d, name, value):
    try:
        s = Select(d.find_element(By.NAME, name))
        s.select_by_value(value)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def overwrite_file():
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "users_not_found.txt"), "w") as myfile:
            myfile.write("userId,A,name\n")
            return False
    except Exception as e:
        print(e)
        return True
        pass

def write_file_data(data):
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "users_not_found.txt"), "a") as myfile:
            myfile.write(data)
            return False
    except Exception as e:
        print(e)
        return True
        pass

def get_vaccine_by_name(vac_name, imm_list):
    for i in range(len(imm_list)):
        if imm_list[i]['Option'].lower() == vac_name.lower():
            return imm_list[i]

def clean_text(txt):
    try:
        txt = txt.replace(".0", "")
        new_string = ''.join(char for char in txt if char.isalnum())
        return new_string
    except:
        pass

def getData(obj , name):
  try:
    return obj[name]
  except Exception as e:
    print(e)
    return ""

def selectSheet(sheet):
    global selected_sheet
    selected_sheet = sheet
    print(selected_sheet)

def sendRequest(subject, message, error = True):
    try:
        payload = {
            "computer": socket.gethostname(),
            "subject": subject,
            "message": message,
            "error": error,
            "bot": "pcc"
        }
        print(payload)
        r = requests.post("https://2qpxr842pk.execute-api.us-east-1.amazonaws.com/Prod/post-sns-data", data=json.dumps(payload))
        return payload
    except Exception as e:
        print(e)
        return ""
def main_loop():
    # read excel
    global driver,patient_data_sheet, current_users, user_name, password, selected_sheet

    print("SELECTED SHEET " + selected_sheet)


    # Read Immunizations.xlsx Excel file
    df = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx"), sheet_name=selected_sheet)
    imm_list = []
    for index, row in df.iterrows():
        imm_list.append(row)

    file_immunizations = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Immunizations.xlsx"), sheet_name='Sheet1')
    immunizations_list = []
    for index, row in file_immunizations.iterrows():
        immunizations_list.append(row)

    # Loop through patient list
    for data in imm_list:
        try:
            driver.refresh()
            t.sleep(2)
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            print("date and time =", dt_string, '>>>>logging in  ')	

            target_user_a = clean_text(str(getData(data,'A#')))
            target_user_id = clean_text(str(getData(data,'userId')))
            targent_name = getData(data,'First Name') + ' ' + getData(data,'Last Name')

            print("Processing - " + target_user_id + "-" + target_user_a + "-" + targent_name)
            select_window(driver, 0)
            t.sleep(1)
            select_window(driver, 0)
            res = select_window(driver, 0)
            if res:
                sendRequest(targent_name, "Error: Could not switch to main window", True)
                break
            # overwrite_file()
            # Stop If Stop Button is pressed
            if thread_stopped == True:
                break

            t.sleep(5)

            # Click on search input field
            try:
                search_select=driver.find_element(By.ID, 'searchSelect')
                ActionChains(driver).move_to_element(search_select).click(search_select).perform()
            except:
                pass

            t.sleep(2)
            try:
                driver.find_element(By.XPATH, '//td[contains(text(),"All Facilities")]').click()
            except:
                pass
            
            # Click on search input field
            # Send Excel name
            res = send_text(driver, 'searchField', target_user_id)
            if res:
                sendRequest(targent_name, "Error: Unable to add text to searchField", True)
                break
            res = send_enter(driver, 'searchField')
            if res:
                sendRequest(targent_name, "Error: Unable to send enter key to searchField", True)
                break
            
            # Stop If Stop Button is pressed
            if thread_stopped == True:
                break

            # Select pop-up window "Global Resident Search -- All Residents"
            t.sleep(5)
            select_window(driver, -1)
            t.sleep(1)
            print("Select pop-up window - Global Resident Search")
            res = select_window(driver, -1)
            if res:
                sendRequest(targent_name, "Error: Could not switch pop-up window", True)
                break

            try:
                driver.find_element(By.XPATH, '//a[contains(text(),"Current"]').click()
                print("Click Current tab")
            except:
                pass

            try:
                driver.find_element(By.XPATH, "//a[@href='/admin/client/clientlist.jsp?ESOLview=Current&amp;ESOLglobalclientsearch=Y']").click();
                print("Click Current tab2")
            except:
                pass

            try:
                driver.find_element(By.LINK_TEXT, "Current").click()
                print("Click Current tab3")
            except:
                pass

            try:
                print("Get Found Users Count")
                current_users=driver.find_element(By.CLASS_NAME, 'pccTableShowDivider').find_elements(By.TAG_NAME, 'tr')
                users_len=len(current_users)
                print(users_len)
                print(current_users[1].text)
                if users_len > 1 and current_users[1].text != "No records found.":
                    for i in range(1, users_len):
                        print("Click on user name")
                        try:
                            user=current_users[i].find_elements(By.TAG_NAME, 'td')[1].find_element(By.TAG_NAME, 'a')
                            ActionChains(driver).move_to_element(user).click(user).perform()
                            t.sleep(5)
                        except:
                            # sendRequest(targent_name, "Error: Unable to click user name", True)
                            break

                        try:
                            driver.switch_to.window(driver.window_handles[0])
                            print("Switch Main Screen")
                        except:
                            break

                        try:
                            driver.switch_to_window(driver.window_handles[0])
                            print("Switch Main Scree1")
                        except:
                            sendRequest(targent_name, "Error: Unable to switch to main screen", True)
                            pass
                        t.sleep(2)
                        try:
                            print("Click on edit button")
                            driver.find_element(By.XPATH, '//span[contains(text(),"Edit")]').click()
                        except:
                            sendRequest(targent_name, "Error: Unable to click on edit button", True)
                            break

                        t.sleep(2)
                        try:
                            print("Demographics")
                            driver.find_element(By.XPATH, '//a[contains(text(),"Demographics")]').click()
                        except:
                            pass
                        
                        try:
                            print("Demographics2")
                            driver.find_element(By.XPATH, "//a[@href='javascript:editDemographicInfo('nonAdmin',3632088)']").click();
                        except:
                            pass

                        t.sleep(3)
                        try:
                            print("Get User A #")
                            user_a=driver.find_elements(By.NAME, 'clientids')[4]
                            user_a_value= user_a.get_attribute("value")
                            print(user_a_value)
                            #print(user_a.text )
                            if user_a_value == target_user_a:#213139072-220786141
                                print("Found target A#")
                                # Click On Cancel button
                                try:
                                    driver.find_element(By.XPATH, "//input[@value='Cancel']").click()
                                except:
                                    pass
                                t.sleep(3)
                                ## START - USER EDIT
                                
                                vaccines_list = []
                                if clean_text(getData(data,'Influenza')).lower() == 'yes':
                                    print('Add Influenza')
                                    vaccines_list.append('Influenza')
                                if clean_text(getData(data,'Tdap')).lower() == 'yes':
                                    print('Add Tdap')
                                    vaccines_list.append('Tdap')
                                if clean_text(getData(data,'Td')).lower() == 'yes':
                                    print('Add Td')
                                    vaccines_list.append('Td')
                                if clean_text(getData(data,'Hepatitis A')).lower() == 'yes':
                                    print('Add Hepatitis A')
                                    vaccines_list.append('Hepatitis A')
                                if clean_text(getData(data,'Hepatitis B')).lower() == 'yes':
                                    print('Add Hepatitis B')
                                    vaccines_list.append('Hepatitis B')
                                if clean_text(getData(data,'HPV')).lower() == 'yes':
                                    print('Add HPV')
                                    vaccines_list.append('HPV')
                                if clean_text(getData(data,'IPV')).lower() == 'yes':
                                    print('Add IPV')
                                    vaccines_list.append('IPV')
                                if clean_text(getData(data,'Meningicoccal')).lower() == 'yes':
                                    print('Add Meningicoccal')
                                    vaccines_list.append('Meningicoccal')
                                if clean_text(getData(data,'MMR')).lower() == 'yes':
                                    print('Add MMR')
                                    vaccines_list.append('MMR')
                                if clean_text(getData(data,'Varicella')).lower() == 'yes':
                                    print('Add Varicella')
                                    vaccines_list.append('Varicella')
                                if clean_text(getData(data,'SARS-COV-2')).lower() == 'yes':
                                    print('Add SARS-COV-2')
                                    vaccines_list.append('SARS-COV-2')
                                if clean_text(getData(data,'SARS-COV-2 < 12')).lower() == 'yes':
                                    print('Add SARS-COV-2 < 12')
                                    vaccines_list.append('SARS-COV-2 < 12')

                                if len(vaccines_list) > 0:
                                    print("Add immunizations")
                                    # Click on Immun tab
                                    res = click_link(driver, "Immun")
                                    if res:
                                        # sendRequest(targent_name, "Error: Unable to click on immunizations", True)
                                        pass
                                    try:
                                        immun=driver.find_element(By.XPATH, '/html/body/table[6]/tbody/tr[2]/td/ul/li[6]/a')
                                        ActionChains(driver).move_to_element(immun).click(immun).perform()
                                    except:
                                        # sendRequest(targent_name, "Error: Unable to click on immunizations table", True)
                                        pass
                                    t.sleep(2)
                                    # Click on New button
                                    res = click_button_value(driver, "New")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click on new button", True)
                                        break
                                    t.sleep(2)
                                    wait_window(driver)
                                    # Select popup window
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                        break

                                    for vacc in vaccines_list:
                                        wait_button(driver, "cancelButton", By.ID)
                                        print("GET immunizations_list " + vacc)
                                        select_vacc = get_vaccine_by_name(vacc, immunizations_list)

                                        # Menu select "Immunization"
                                        t.sleep(1)
                                        print(select_vacc["Name"] + "-" + clean_text(str(select_vacc["Menu val"])))
                                        res = select_menu_name(driver, "immunizationId",clean_text( str(select_vacc["Menu val"])))
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to select menu " + select_vacc["Name"], True)
                                            break
                                        if select_vacc["Search"] == "Yes":
                                            res = click_link_href(driver, "javascript:cvxCodeSearch();")
                                            if res:
                                                sendRequest(targent_name, "Error: Unable to click on search button", True)
                                                break
                                            t.sleep(0.5)
                                            wait_window(driver)
                                            # Select popup window
                                            select_window(driver, -1)
                                            t.sleep(1)
                                            res = select_window(driver, -1)
                                            if res:
                                                sendRequest(targent_name, "Error: Unable to select window", True)
                                                break
                                            send_text(driver, "searchText", select_vacc["Search Name"])
                                            send_enter(driver, "searchText")
                                            t.sleep(1)
                                            vac_res =driver.find_element(By.CLASS_NAME, 'pccResults').find_elements(By.TAG_NAME, 'tr')
                                            if len(vac_res[1].find_elements(By.TAG_NAME, 'a')) > 0:
                                                vac_res[int(select_vacc["Search Pos"])].find_elements(By.TAG_NAME, 'a')[0].click()
                                                t.sleep(0.5)
                                                select_window(driver, 0)
                                                t.sleep(1)
                                                select_window(driver, 0)
                                                t.sleep(0.5)
                                                select_window(driver, -1)
                                                t.sleep(1)
                                                select_window(driver, -1)
                                        # Menu select "Given"
                                        res = select_menu_name(driver, "consentGiven", "Y")
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to select menu Given", True)
                                            break
                                        t.sleep(0.5)
                                        # Set date of visit
                                        res = send_text(driver, "dateGiven_dummy", get_string_date(getData(data,'Date of visit')))
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to set date of visit", True)
                                            break
                                        # Set notes
                                        notes =  select_vacc["Name"] + "\n" + select_vacc["Zone"] + "\n" + "Lot# " + str(select_vacc["Lot#"]) + "\n" + "Exp: " + get_string_date(select_vacc["Exp"]) + "\n" + "Manufacturer: " + select_vacc["Manufacturer"] + "\n" + "VIS Date: " + get_string_date(vis_date) + "\n" + "VIS Given: " + get_string_date(getData(data,'Date of visit')) + "\n" + "Funding: " + select_vacc["Funding"]
                                        res = send_text_name(driver, "notes", notes)
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to set notes", True)
                                            break
                                        # Click on button "Save & New"
                                        res = click_button_value(driver, "Save & New")
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to click on save and new button", True)
                                            break
                                        t.sleep(2)
                                    t.sleep(1)
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                        break
                                    wait_button(driver, "cancelButton", By.ID)
                                    send_click(driver, "cancelButton")
                                    # Todo: Marco Martinez - Click on cancel button

                                if clean_text(getData(data,'Dose 2 SARS-COV-2')).lower() == 'yes' or clean_text(getData(data,'Dose 2 SARS-COV-2 < 12')).lower() == 'yes':
                                    print("Add SARS-COV-2 Dose 2")
                                    # Click on Immun tab
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                        break
                                    t.sleep(1)
                                    res = click_link(driver, "Immun")
                                    if res:
                                        # sendRequest(targent_name, "Error: Unable to click on immunizations link", True)
                                        pass
                                    try:
                                        immun=driver.find_element(By.XPATH, '/html/body/table[6]/tbody/tr[2]/td/ul/li[6]/a')
                                        ActionChains(driver).move_to_element(immun).click(immun).perform()
                                    except:
                                        # sendRequest(targent_name, "Error: Unable to click on immunizations table", True)
                                        pass
                                    t.sleep(3)
                                    res = send_click_pos_by_class(driver, "listbuttonred", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click on listbuttonred", True)
                                        pass
                                    t.sleep(2)
                                    wait_window(driver)
                                    # Select popup window
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                        break
                                    vaccines_list_dose = []
                                    if clean_text(getData(data,'Dose 2 SARS-COV-2')).lower() == 'yes':
                                        vaccines_list_dose.append('SARS-COV-2')
                                    if clean_text(getData(data,'Dose 2 SARS-COV-2 < 12')).lower() == 'yes':
                                        vaccines_list_dose.append('SARS-COV-2 < 12')
                                    
                                    for vacc in vaccines_list_dose:
                                        wait_button(driver, "cancelButton", By.ID)
                                        select_vacc = get_vaccine_by_name(vacc, immunizations_list)
                                        # Menu select "Immunization"
                                        t.sleep(1)
                                        print(select_vacc["Name"] + "-" + clean_text(str(select_vacc["Menu val"])))
                                        if select_vacc["Search"] == "Yes":
                                            res = click_link_href(driver, "javascript:cvxCodeSearch();")
                                            if res:
                                                sendRequest(targent_name, "Error: Unable to click on search link", True)
                                                break
                                            t.sleep(0.5)
                                            wait_window(driver)
                                            # Select popup window
                                            select_window(driver, -1)
                                            t.sleep(1)
                                            res = select_window(driver, -1)
                                            if res:
                                                sendRequest(targent_name, "Error: Unable to select window", True)
                                                break
                                            res = send_text(driver, "searchText", select_vacc["Search Name"])
                                            if res:
                                                sendRequest(targent_name, "Error: Unable to set search text", True)
                                                break
                                            res = send_enter(driver, "searchText")
                                            if res:
                                                sendRequest(targent_name, "Error: Unable to send enter", True)
                                                break
                                            t.sleep(1)
                                            vac_res =driver.find_element(By.CLASS_NAME, 'pccResults').find_elements(By.TAG_NAME, 'tr')
                                            if len(vac_res[1].find_elements(By.TAG_NAME, 'a')) > 0:
                                                vac_res[int(select_vacc["Search Pos"])].find_elements(By.TAG_NAME, 'a')[0].click()
                                                t.sleep(0.5)
                                                select_window(driver, 0)
                                                t.sleep(1)
                                                select_window(driver, 0)
                                                t.sleep(0.5)
                                                select_window(driver, -1)
                                                t.sleep(1)
                                                select_window(driver, -1)
                                        # Menu select "Given"
                                        res = select_menu_name(driver, "consentGiven", "Y")
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to select menu consentGiven", True)
                                            break
                                        t.sleep(0.5)
                                        # Set date of visit
                                        res = send_text(driver, "dateGiven_dummy", get_string_date(getData(data,'Date of visit')))
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to set date of visit", True)
                                            break
                                        # Set notes
                                        notes =  select_vacc["Name"]  + "\n" + select_vacc["Zone"] + "\n" + "Lot# " + str(select_vacc["Lot#"]) + "\n" + "Exp: " + get_string_date(select_vacc["Exp"]) + "\n" + "Manufacturer: " + select_vacc["Manufacturer"] + "\n" + "VIS Date: " + get_string_date(vis_date) + "\n" + "VIS Given: " + get_string_date(getData(data,'Date of visit')) + "\n" + "Funding: " + select_vacc["Funding"]
                                        res = send_text_name(driver, "notes", notes)
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to set notes", True)
                                            break
                                        # Click on button "Save & New"
                                        res = click_button_value(driver, "Save & New")
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to click on save & new button", True)
                                            break
                                        t.sleep(2)
                                    t.sleep(1)
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                        break
                                    wait_button(driver, "cancelButton", By.ID)
                                    send_click(driver, "cancelButton")
                                    # Todo: Marco Martinez - Click on cancel button
                                    
                                if clean_text(getData(data,'Initial Medical Form')).lower() == 'yes':
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                        break
                                    print("*Initial Medical Exam Unaccompanied Children's Program Office of Refugee Resettlement (ORR) - V4 ")
                                    # Click on "Assmnts"
                                    try:
                                        assessment=driver.find_element(By.XPATH, '/html/body/table[6]/tbody/tr[2]/td/ul/li[9]/a')
                                        ActionChains(driver).move_to_element(assessment).click(assessment).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on Assmnts", True)
                                    t.sleep(5)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    # Click on "New" button
                                    try:
                                        newwww= driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td/table/tbody/tr/td[1]/input')
                                        ActionChains(driver).move_to_element(newwww).click(newwww).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on new button", True)
                                        break

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(5)
                                    # Select pop-up window "Reasons for Assessment"
                                    driver.switch_to.window(driver.window_handles[-1])

                                    # Select Assesment
                                    try:
                                        assessment = Select(driver.find_element(By.ID, 'std_assessment'))
                                        # *Initial Medical Exam Unaccompanied Children's Program Office of Refugee Resettlement (ORR)  - V 3 
                                        assessment.select_by_value(value_initial_medical_exam)
                                    except:
                                        sendRequest(targent_name, "Error: Unable to select assessment", True)
                                        pass
                                    # Click on "save" button - value="Save"
                                    try:
                                        driver.find_element(By.XPATH, "//input[@value='Save']").click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on save button", True)
                                        pass

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(3)
                                    # Return to main window
                                    driver.switch_to.window(driver.window_handles[0])
                                    t.sleep(3)
                                    now = datetime.now()
                                    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
                                    print("date and time =[", dt_string, ']>>>>Filling forms  ')	

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section A - General Information")
                                    try:
                                        # A.General Information - a.Name and Designation
                                        driver.find_element(By.ID, 'linkCust_A_1_1').clear()
                                        driver.find_element(By.ID, 'linkCust_A_1_1').send_keys(name_designation)
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set name_designation", True)
                                        pass
                                    try:
                                        driver.find_element(By.ID, 'linkCust_A_1_2').clear()
                                        driver.find_element(By.ID, 'linkCust_A_1_2').send_keys(str(getData(data,'telephone')))
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set telephone", True)
                                        pass
                                    try:
                                        driver.find_element(By.ID, 'linkCust_A_3').clear()
                                        driver.find_element(By.ID, 'linkCust_A_3').send_keys(getData(data,'clinicname') or clinic_or_practice)
                                    except Exception as e:
                                        print(e)
                                        sendRequest(targent_name, "Error: Unable to set clinicname", True)
                                        pass
                                    try:
                                        driver.find_element(By.ID, 'linkCust_A_4').clear()
                                        driver.find_element(By.ID, 'linkCust_A_4').send_keys(getData(data,'Healthcare Provider Street address, City or Town, State'))
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set Healthcare Provider Address", True)
                                        pass
                                    try:
                                        driver.find_element(By.ID, 'linkCust_A_5_dummy').clear()
                                        driver.find_element(By.ID, 'linkCust_A_5_dummy').send_keys(get_string_date(getData(data,'Date of visit')))
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set Date of visit", True)
                                        pass
                                    try:
                                        driver.find_element(By.ID, 'linkCust_A_7').clear()
                                        driver.find_element(By.ID, 'linkCust_A_7').send_keys(getData(data,'Program Name'))
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set Program Name", True)
                                        pass
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section B - History and Physical")
                                    # B.a History and Physical - Allergies 
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_B_2')[0].click()                     
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on Allergies", True)
                                        pass
                                    #  B.-4
                                    try:
                                        driver.find_element(By.ID, 'linkCust_B_4').clear()
                                        driver.find_element(By.ID, 'linkCust_B_4').send_keys("20/20")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set B-4", True)
                                        pass
                                    #  B.-4a
                                    try:
                                        driver.find_element(By.ID, 'linkCust_B_4a').clear()
                                        driver.find_element(By.ID, 'linkCust_B_4a').send_keys("20/20")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set B-4a", True)
                                        pass
                                    #  B.-4b
                                    try:
                                        driver.find_element(By.ID, 'linkCust_B_4b').clear()
                                        driver.find_element(By.ID, 'linkCust_B_4b').send_keys("20/20")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set B-4b", True)
                                        pass
                                    #  B.-4c (radio button)  - a. 	Pass
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_B_4c')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on B-4c", True)
                                        pass
                                    #  B.-5 (radio button)  - a. - Yes
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_B_5')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on B-5", True)
                                        pass
                                    #  B.-6 (radio button)  - a. - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_B_6')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on B-6", True)
                                        pass
                                    #  B.-7
                                    try:
                                        driver.find_element(By.ID, 'linkCust_B_7').clear()
                                        driver.find_element(By.ID, 'linkCust_B_7').send_keys("Denies")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set B-7", True)
                                        pass
                                    #  B.-8 (radio button)  - a. - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_B_8')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on B-8", True)
                                        pass
                                    #  B.-9
                                    try:
                                        driver.find_element(By.ID, 'linkCust_B_9').clear()
                                        driver.find_element(By.ID, 'linkCust_B_9').send_keys("Denies")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set B-9", True)
                                        pass
                                    #  B.-10
                                    try:
                                        driver.find_element(By.ID, 'linkCust_B_10').clear()
                                        driver.find_element(By.ID, 'linkCust_B_10').send_keys("The child reports traveling from")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set B-10", True)
                                        pass
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section C - Review of Systems")
                                    # C.1 - No abnormal Findings
                                    try:
                                        driver.find_element(By.ID, 'linkCust_C_1').click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on C.1", True)
                                        pass
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section D - Physical Examination")
                                    # D.1 - General Appearance	- Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_1')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.1", True)
                                        pass
                                    # D.2 - heent - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_2')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.2", True)
                                        pass
                                    # D.3 - Neck - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_3')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.3", True)
                                        pass
                                    # D.4 - Heart - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_4')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.4", True)
                                        pass
                                    # D.5 - Lungs - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_5')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.5", True)
                                        pass
                                    # D.6 - Abdomen - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_6')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.6", True)
                                        pass
                                    # D.8 - Extremities - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_8')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.8", True)
                                        pass
                                    # D.9 - Back/Spine	 - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_9')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.9", True)
                                        pass
                                    # D.10 - Neurologic	 - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_10')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.10", True)
                                        pass
                                    # D.11 - Neurologic	 - Normal
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_D_11')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on D.11", True)
                                        pass
                                    # D.12 - Other
                                    try:
                                        driver.find_element(By.ID, 'linkCust_D_12').clear()
                                        driver.find_element(By.ID, 'linkCust_D_12').send_keys("Whisper test passed")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set D.12", True)
                                        pass
                                    
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section E - Psychosocial Risk")
                                    # E.1 - Mental Health concerns ( ≤3 mos) - Denied, with no obvious sign/symptoms
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_E_1')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on E.1", True)
                                        pass
                                    # E.2 - Mental Health concerns ( ≤3 mos) - Denied, with no obvious sign/symptoms
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_E_2')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on E.2", True)
                                        pass
                                    # E.3 - Mental Health concerns ( ≤3 mos) - Denied, with no obvious sign/symptoms
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_E_3')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on E.3", True)
                                        pass
                                    # E.4a - Nonconsensual Sexual Activity
                                    try:
                                        driver.find_element(By.ID, 'linkCust_E_4a').clear()
                                        driver.find_element(By.ID, 'linkCust_E_4a').send_keys("Denied.")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set E.4a", True)
                                        pass
                                    
                                    # E.5 - Mental Health concerns ( ≤3 mos) - Denied, with no obvious sign/symptoms
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_E_5')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on E.5", True)
                                        pass
                                    
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section F - 	Laboratory Testing")
                                    # F.5 - HIV
                                    try:
                                        driver.find_element(By.ID, 'linkCust_F_5').click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on F.5", True)
                                        pass
                                    t.sleep(1)
                                    # F.5-a - > 13 yrs or sexual activity - Negative
                                    try:
                                        driver.find_element(By.ID, 'linkCust_F_5a').click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on F.5a", True)
                                        pass
                                    # F.5-b -Test: Rapid Oral
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_F_5b')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on F.5b", True)
                                        pass
                                    print("Section G - TB Screening")
                                    # G.1 - a - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_G_1')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on G.1", True)
                                        pass
                                    # G.2 - a - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_G_2')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on G.2", True)
                                        pass
                                    # G.3 - a - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_G_3')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on G.3", True)
                                        pass
                                    # G.4 - b - IGRA (<2yrs)
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_G_4')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on G.4", True)
                                        pass
                                    # G.4 - c - CXR (<15yrs)
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_G_4')[2].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on G.4", True)
                                        pass

                                    print("Section H - Diagnosis and Plan")
                                    # H.1 - b - Yes
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_H_1')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on H.1", True)
                                        pass
                                    t.sleep(1)
                                    # # H.13 - Medical Order
                                    # try:
                                    #     driver.find_element(By.ID, 'linkCust_H_13').clear()
                                    #     driver.find_element(By.ID, 'linkCust_H_13').send_keys("Encounter for screening for other viral diseases")
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to set H.13", True)
                                    #     pass
                                    # H.a - a - 	Return to clinic- PRN/As needed
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_H_a')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on H.a", True)
                                        pass
                                    # H.b-a - Yes
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_H_b')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on H.b", True)
                                        pass
                                    t.sleep(1)
                                    # H.b1
                                    try:
                                        driver.find_element(By.ID, 'linkCust_H_b1').clear()
                                        driver.find_element(By.ID, 'linkCust_H_b1').send_keys("The minor is medically cleared to travel if no known exposure to COVID has occurred and no other concerns requiring medical follow-up and/or specialty follow-up have been identified in subsequent visits.")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set H.b1", True)
                                        pass

                                    # H.c-a - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_H_c')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on H.c", True)
                                        pass
                                    # H.d-a - No
                                    try:
                                        driver.find_elements(By.ID, 'linkCust_H_d')[0].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on H.d", True)
                                        pass
                                    # H.f - Minor requires quarantine/ isolation, specify diagnosis and timeframe:
                                    # try:
                                    #     driver.find_element(By.ID, 'linkCust_H_f').clear()
                                    #     driver.find_element(By.ID, 'linkCust_H_f').send_keys("Quarantine minor for 7 days from day of arrival.")
                                    # except:
                                    #     pass
                                    # H.g - Immunizations given/validated from foreign record
                                    try:
                                        driver.find_element(By.ID, 'linkCust_H_g').click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on H.g", True)
                                        pass
                                    # # H.i - Age- appropriate anticipatory guidance discussed and/or handout given
                                    # try:
                                    #     driver.find_element(By.ID, 'linkCust_H_i').click()
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on H.i", True)
                                    #     pass
                                    # H.p - Recommendations from Healthcare Provider/ Additional Information:
                                    try:
                                        driver.find_element(By.ID, 'linkCust_H_p').clear()
                                        driver.find_element(By.ID, 'linkCust_H_p').send_keys("The minor is medically cleared to travel if no known exposure to COVID has occurred and no other concerns requiring medical follow-up and/or specialty follow-up have been identified in subsequent visits.\n\n Scribed by:")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to set H.p", True)
                                        pass
                                    print("Click on button Save & Exit")
                                    try:
                                        driver.find_elements(By.ID, 'saveandexitbutton')[1].click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on button Save & Exit", True)
                                        pass
                                    t.sleep(3)                            
                                if clean_text(getData(data,'Quarantine Form')).lower() == 'yes':
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                    print("Quarantine/Isolation")
                                    # Click on "Assmnts"
                                    try:
                                        assessment=driver.find_element(By.XPATH, '/html/body/table[6]/tbody/tr[2]/td/ul/li[9]/a')
                                        ActionChains(driver).move_to_element(assessment).click(assessment).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on Assmnts", True)
                                        pass
                                    t.sleep(5)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    # Click on "New" button
                                    try:
                                        newwww= driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td/table/tbody/tr/td[1]/input')
                                        ActionChains(driver).move_to_element(newwww).click(newwww).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on New", True)
                                        pass
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(5)
                                    # Select pop-up window "Reasons for Assessment"
                                    driver.switch_to.window(driver.window_handles[-1])

                                    try:
                                        # Select Assesment
                                        assessment = Select(driver.find_element(By.ID, 'std_assessment'))
                                        # Quarantine/Isolation 
                                        assessment.select_by_value(value_quarantine_isolation)
                                    except:
                                        sendRequest(targent_name, "Error: Unable to select Assesment", True)
                                        pass

                                    try:
                                        # Click on "save" button - value="Save"
                                        driver.find_element(By.XPATH, "//input[@value='Save']").click()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on save", True)
                                        pass

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(3)
                                    # Return to main window
                                    driver.switch_to.window(driver.window_handles[0])
                                    t.sleep(3)	

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section A - Demographics")
                                    # A.1.
                                    res = send_text(driver, "linkCust_A_1", name_designation)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.1", True)
                                    # A.1.a.
                                    res = send_text(driver, "linkCust_A_1a_dummy", get_string_date(getData(data,'Date of visit')))
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.1.a", True)
                                    # A.2.b.
                                    res = send_click_pos(driver, "linkCust_A_2", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.2.b", True)
                                    # A.3.e.
                                    res = send_click_pos(driver, "linkCust_A_3", 4)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.3.e", True)
                                    t.sleep(1)
                                    # A.3.a.
                                    res = send_text(driver, "linkCust_A_3a", "Contact with and (suspected) exposure to COVID-19")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.3.a", True)
                                    # B.6.d.
                                    res = send_click_pos(driver, "linkCust_B_6", 3)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set B.6.d", True)
                                    t.sleep(1)
                                    # B.6.a
                                    res = send_text(driver, "linkCust_B_6a", "Quarantine minor for 7 days from day of arrival")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set B.6.a", True)
                                    # C.7.b.
                                    res = send_click_pos(driver, "linkCust_C_7", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set C.7.b", True)
                                    t.sleep(1)
                                    # C.9.
                                    res = send_click(driver, "linkCust_C_9")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set C.9", True)
                                    # C.9.a.
                                    res = send_click(driver, "linkCust_C_9a")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set C.9.a", True)
                                    # C.9.b.
                                    res = send_click(driver, "linkCust_C_9b")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set C.9.b", True)
                                    # Click Save & Sign & Lock & Exit
                                    # send_click_pos(driver, "saveandexitbutton", 1)
                                    # t.sleep(5)

                                    res = send_click_pos(driver, "saveandsignbutton", 3)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click saveandsignbutton", True)
                                    t.sleep(5)
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                    t.sleep(1)
                                    res = send_text_name(driver, "pw", password)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set password", True)
                                    res = send_click(driver, "saveButton")
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click saveButton", True)
                                    print("Standing Orders Form Completed")
                                    t.sleep(3)
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)

                                if clean_text(getData(data,'Standing Orders Form')).lower() == 'yes':
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to select window", True)
                                    print("Standing Orders 12 and Over")
                                    # Click on "Assmnts"
                                    try:
                                        assessment=driver.find_element(By.XPATH, '/html/body/table[6]/tbody/tr[2]/td/ul/li[9]/a')
                                        ActionChains(driver).move_to_element(assessment).click(assessment).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on Assmnts", True)
                                        pass
                                    t.sleep(5)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    # Click on "New" button
                                    try:
                                        newwww= driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td/table/tbody/tr/td[1]/input')
                                        ActionChains(driver).move_to_element(newwww).click(newwww).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on New", True)
                                        pass
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(5)
                                    # Select pop-up window "Reasons for Assessment"
                                    driver.switch_to.window(driver.window_handles[-1])

                                    try:
                                        # Select Assesment
                                        assessment = Select(driver.find_element(By.ID, 'std_assessment'))
                                        # Standing Orders 12 and Over 
                                        assessment.select_by_value(value_standing_orders_12_over)
                                    except:
                                        sendRequest(targent_name, "Error: Unable to select Assesment", True)
                                        pass

                                    # Click on "save" button - value="Save"
                                    driver.find_element(By.XPATH, "//input[@value='Save']").click()

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(3)
                                    # Return to main window
                                    driver.switch_to.window(driver.window_handles[0])
                                    t.sleep(3)	

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    print("Section A - Standing Orders 12yr and older")
                                    # A.1.a.
                                    res = send_click_pos(driver, "linkCust_A_1", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.1.a", True)
                                    # A.1.b.
                                    res = send_click_pos(driver, "linkCust_A_1", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.1.b", True)
                                    # A.2.a.
                                    res = send_click_pos(driver, "linkCust_A_2", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.2.a", True)
                                    # A.2.b.
                                    res = send_click_pos(driver, "linkCust_A_2", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.2.b", True)
                                    # A.2.c.
                                    res = send_click_pos(driver, "linkCust_A_2", 2)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.2.c", True)
                                    # A.3.a.
                                    res = send_click_pos(driver, "linkCust_A_3", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.3.a", True)
                                    res = send_click_pos(driver, "linkCust_A_3", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.3.b", True)
                                    # A.4.a.
                                    res = send_click_pos(driver, "linkCust_A_4", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.4.a", True)
                                    # A.5.a.
                                    res = send_click_pos(driver, "linkCust_A_5", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.5.a", True)
                                    # A.6.a.
                                    res = send_click_pos(driver, "linkCust_A_6", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.6.a", True)
                                    # A.7.a.
                                    res = send_click_pos(driver, "linkCust_A_7", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.7.a", True)
                                    res = send_click_pos(driver, "linkCust_A_7", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.7.b", True)
                                    # A.8.a.
                                    res = send_click_pos(driver, "linkCust_A_8", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.8.a", True)
                                    res = send_click_pos(driver, "linkCust_A_8", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.8.b", True)
                                    # A.9.a.
                                    res = send_click_pos(driver, "linkCust_A_9", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.9.a", True)
                                    # A.10.a.
                                    res = send_click_pos(driver, "linkCust_A_10", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.10.a", True)
                                    # A.11.a.
                                    res = send_click_pos(driver, "linkCust_A_11", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.11.a", True)
                                    # A.12.a.
                                    res = send_click_pos(driver, "linkCust_A_12", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.12.a", True)
                                    # A.13.a.
                                    res = send_click_pos(driver, "linkCust_A_13", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.13.a", True)
                                    # A.14.a.
                                    res = send_click_pos(driver, "linkCust_A_14", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.14.a", True)
                                    res = send_click_pos(driver, "linkCust_A_14", 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.14.b", True)
                                    # A.15.a.
                                    res = send_click_pos(driver, "linkCust_A_15", 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.15.a", True)
                                    # Click Save & Sign & Lock & Exit
                                    # send_click_pos(driver, "saveandexitbutton", 1)
                                    # t.sleep(5)
                                    res = send_click_pos(driver, "saveandsignbutton", 3)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.15.a", True)
                                    t.sleep(5)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.15.a", True)
                                    t.sleep(1)
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = send_text_name(driver, "pw", password)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set password", True)
                                    send_click(driver, "saveButton")
                                    print("Standing Orders Form Completed")
                                    t.sleep(3)
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to switch window", True)
                                if clean_text(getData(data,'Assessment')).lower() == 'yes':
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to switch window", True)
                                    print("*Health Assessment Form Unaccompanied Children's Program Office of Refugee Resettlement (ORR) - V 2")
                                    # Click on "Assmnts"
                                    try:
                                        assessment=driver.find_element_by_xpath('/html/body/table[6]/tbody/tr[2]/td/ul/li[9]/a')
                                        ActionChains(driver).move_to_element(assessment).click(assessment).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on Assmnts", True)
                                    t.sleep(10)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    try:
                                        # Click on "New" button
                                        newwww= driver.find_element_by_xpath('/html/body/form/table/tbody/tr[1]/td/table/tbody/tr/td[1]/input')
                                        ActionChains(driver).move_to_element(newwww).click(newwww).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on New", True)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(5)
                                    # Select pop-up window "Reasons for Assessment"
                                    # Click on "save" button - value="Save"
                                    select_window(driver, -1)
                                    t.sleep(1)
                                    res = select_window(driver, -1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to switch window", True)

                                    try:
                                        saveeee=driver.find_element_by_xpath('/html/body/form/div[2]/input[1]')
                                        ActionChains(driver).move_to_element(saveeee).click(saveeee).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on save", True)
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    t.sleep(5)
                                    # Return to main window
                                    select_window(driver, 0)
                                    t.sleep(1)
                                    res = select_window(driver, 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to switch window", True)
                                    t.sleep(1)

                                    # A.General Information - a.Name and Designation
                                    res = send_text(driver, 'linkCust_A_1_1', name_designation)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.1.1", True)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    res = send_text(driver, 'linkCust_A_1_2', str(getData(data,'telephone')))
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.1.2", True)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break
                                    res = send_text(driver, 'linkCust_A_2', md_do_pa_np)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.2", True)
                                    res = send_text(driver, 'linkCust_A_3', getData(data,'clinicname') or clinic_or_practice)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.3", True)
                                    res = send_text(driver, 'linkCust_A_4', getData(data,'Healthcare Provider Street address, City or Town, State'))
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.4", True)
                                    # send_text(driver, 'linkCust_A_5_dummy', data['Date of visit'])

                                    res = send_text(driver, "linkCust_A_5_dummy", get_string_date(getData(data,'Date of visit')))
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set A.5", True)
                                    # send_text(driver, 'linkCust_A_7', data['Program Name'])
                                    # send_text(driver, 'Cust_B_1', temp_c)
                                    # send_text(driver, 'linkCust_B_1a', hr)
                                    # send_text(driver, 'linkCust_B_1b', bp)
                                    # send_text(driver, 'linkCust_B_1c', rr)
                                    # send_text(driver, 'linkCust_B_1d', ht)
                                    # send_text(driver, 'linkCust_B_1e', wt)

                                    #B.1a. Clear Data
                                    # try:
                                    #     driver.find_element(By.XPATH, '//a[@href="javascript:clearPPControl(\'linkCust_B_1a\');"]').click();
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on B.1a", True)
                                    #     pass
                                    # #B.1b. Clear Data
                                    # try:
                                    #     driver.find_element(By.XPATH, '//a[@href="javascript:clearPPControl(\'linkCust_B_1b\');"]').click();
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on B.1b", True)
                                    #     pass
                                    # #B.1c. Clear Data
                                    # try:
                                    #     driver.find_element(By.XPATH, '//a[@href="javascript:clearPPControl(\'linkCust_B_1c\');"]').click();
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on B.1c", True)
                                    #     pass
                                    # #B.1d. Clear Data
                                    # try:
                                    #     driver.find_element(By.XPATH, '//a[@href="javascript:clearPPControl(\'linkCust_B_1d\');"]').click();
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on B.1d", True)
                                    #     pass
                                    # #B.1e. Clear Data
                                    # try:
                                    #     driver.find_element(By.XPATH, '//a[@href="javascript:clearPPControl(\'linkCust_B_1e\');"]').click();
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on B.1e", True)
                                    #     pass
                                    # #B.1f. Clear Data
                                    # try:
                                    #     driver.find_element(By.XPATH, '//a[@href="javascript:clearPPControl(\'linkCust_B_1f\');"]').click();
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click on B.1f", True)
                                    #     pass

                                    #  B. History and Physical - Allergies 
                                    if no_allergies == 'Yes':
                                        res = send_click_pos(driver, 'linkCust_B_2', 0)
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to click on B.2", True)
                                    if food == 'Yes':
                                        res = send_click_pos(driver, 'linkCust_D_2', 1)
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to click on D.2", True)
                                    if medication == 'Yes':
                                        res = send_click_pos(driver, 'linkCust_D_2', 2)
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to click on D.2", True)

                                    # B: History and Physical - 3. Concerns expressed by child or caregiver?
                                    # Click Yes
                                    # linkCust_B_3
                                    try:
                                        b_3=driver.find_elements_by_id('linkCust_B_3')
                                        ActionChains(driver).move_to_element( b_3[1]).click( b_3[1]).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click on B.3", True)
                                        pass

                                    res = send_text(driver, 'linkCust_B_3a', corrected_left_eye)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set B.3a", True)

                                    # res = send_text(driver, 'linkCust_B_3b', corrected_both_eyes)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.3b", True)

                                    # res = send_text(driver, 'linkCust_B_4', uncorrected_right_eye)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.4", True)

                                    # res = send_text(driver, 'linkCust_B_4b', uncorrected_both_eyes)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.4b", True)

                                    # res = send_text(driver, 'linkCust_B_5', medical_history)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.5", True)

                                    # res = send_text(driver, 'linkCust_B_6', travel_history)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.6", True)

                                    # res = send_text(driver, 'linkCust_B_7', past_medical_history)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.7", True)

                                    # res = send_text(driver, 'linkCust_B_8', family_history)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.8", True)

                                    # res = send_text(driver, 'linkCust_B_9', str(lmp))
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.9", True)

                                    # res = send_text(driver, 'linkCust_B_9a', str(previous_regnancy))
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to set B.9a", True)



                                    res = send_text(driver, 'linkCust_C_21', other_1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set C.21", True)

                                    res = send_text(driver, 'linkCust_C_22', other_2)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set C.22", True)


                                    if no_abnormal_findings == 'Yes':
                                        res = send_click(driver, 'linkCust_C_1')
                                        if res:
                                            sendRequest(targent_name, "Error: Unable to click on C.1", True)

                                    # Physical Examination D
                                    res = send_click_pos(driver, 'linkCust_D_1', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.1.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_2', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.2.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_3', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.3.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_4', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.4.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_5', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.5.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_6', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.6.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_8', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.8.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_9', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.9.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_10', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.10.a", True)

                                    res = send_click_pos(driver, 'linkCust_D_11', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to set D.11.a", True)


                                    # if general_appearance.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_1', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_1', 1)

                                    # if heent.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_2', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_2', 1)

                                    # if neck.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_3', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_3', 1)

                                    # if heart.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_4', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_4', 1)

                                    # if lungs.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_5', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_5', 1)

                                    # if abdomen.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_6', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_6', 1)
                                        
                                    # try:
                                    #     d7a=driver.find_element_by_id('linkCust_D_7a')
                                    #     ActionChains(driver).move_to_element( d7a).click( d7a).send_keys(describe).perform()
                                    # except:
                                    #     pass

                                    # if extremeties.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_8', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_8', 1)

                                    # if back_spine.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_9', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_9', 1)

                                    # if neurologic.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_10', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_10', 1)

                                    # if skin.lower() == 'normal':
                                    #     send_click_pos(driver, 'linkCust_D_11', 0)
                                    # else:
                                    #     send_click_pos(driver, 'linkCust_D_11', 1)

                                    # try:
                                    #     driver.find_element(By.ID, 'linkCust_D_12').clear()
                                    #     driver.find_element(By.ID, 'linkCust_D_12').send_keys("Whisper test passed")
                                    # except:
                                    #     pass
                                    
                                    # asdddddd= driver.find_element_by_xpath('/html/body/table[5]/tbody/tr[2]/td/table/tbody/tr[5]/td/form/table/tbody/tr[2]/td/table/tbody/tr[5]/td[3]/table/tbody/tr[3]/td/table/tbody/tr[4]/td[2]')

                                    # try:
                                        
                                    #     buttonnn= asdddddd.find_elements_by_tag_name('input')
                                    #     ActionChains(driver).move_to_element(buttonnn[int(mental_health)]).click(buttonnn[int(mental_health)]).perform()
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click mental_health", True)
                                    #     pass

                                    
                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    # # Escondida checkbox
                                    # res = send_click_pos(driver, 'linkCust_F_2', 0)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click F.2", True)

                                    # # Escondida checkbox
                                    # t.sleep(1)
                                    # res = send_click_pos(driver, 'linkCust_F_2b', 1)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click F.2b", True)

                                    # t.sleep(1)
                                    # res = send_click_pos(driver, 'linkCust_G_1', 0)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click G.1", True)

                                    # t.sleep(1)
                                    # res = send_click_pos(driver, 'linkCust_G_2', 0)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click G.2", True)

                                    # res = send_click_pos(driver, 'linkCust_G_3', 0)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click G.3", True)

                                    # res = send_click_pos(driver, 'linkCust_G_4', 1)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click G.4", True)

                                    # # No se encontro
                                    # res = send_click_pos(driver, 'linkCust_H_1', 0)
                                    # if res:
                                    #     sendRequest(targent_name, "Error: Unable to click H.1", True)

                                    # #linkCust_H_13
                                    # try:
                                    #     # No se encontro
                                    #     b1e=driver.find_elements_by_id('linkCust_H_13')
                                    #     ActionChains(driver).move_to_element(b1e[0]).click(b1e[0]).send_keys(other_medical).perform()
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click other_medical", True)
                                    #     pass

                                    # try:
                                    # # t.sleep(5)
                                    #     # No se encontro
                                    #     b1e=driver.find_elements_by_id('linkCust_H_14')
                                    # #ActionChains(driver).move_to_element(b1e[0]).click(b1e[0]).perform()
                                    
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click linkCust_H_14", True)
                                    #     pass
                                    
                                    # try:
                                    # # t.sleep(5)
                                    #     # No se encontro
                                    #     b1e=driver.find_elements_by_id('linkCust_H_15')
                                    #     ActionChains(driver).move_to_element(b1e[0]).click(b1e[0]).perform()
                                    #     ActionChains(driver).move_to_element(b1e[5]).click(b1e[5]).perform()
                                    #     ActionChains(driver).move_to_element(b1e[7]).click(b1e[7]).perform()
                                    #     ActionChains(driver).move_to_element(b1e[10]).click(b1e[10]).perform()
                                        
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click linkCust_H_15", True)
                                    #     pass
                                    # try:
                                    #     # No se encontro
                                    #     b1e=driver.find_elements_by_id('linkCust_H_15e')
                                    #     ActionChains(driver).move_to_element(b1e[0]).click(b1e[0]).send_keys(h15).perform()
                                    # except:
                                    #     sendRequest(targent_name, "Error: Unable to click linkCust_H_15e", True)
                                    #     pass
                                            
                                    # AW: Describe concerns
                                    # try:
                                    #     # Escondido
                                    #     e1a=driver.find_element_by_id('linkCust_E_1a')
                                    #     ActionChains(driver).move_to_element( e1a).click( e1a).send_keys(describe_concerns).perform()
                                    # except:
                                    #     pass
                                    try:
                                        # No se encontro
                                        h16=driver.find_element_by_id('linkCust_H_16')
                                        ActionChains(driver).move_to_element( h16).click( h16).send_keys(getData(data,'Additional Information')).perform()
                                    except:
                                        # sendRequest(targent_name, "Error: Unable to click linkCust_H_16", True)
                                        pass

                                    # No se encontro
                                    send_click_pos(driver, 'linkCust_I_1', 1)
                                    # No se encontro
                                    send_click_pos(driver, 'linkCust_I_2', 2)
                                    # No se encontro
                                    send_click_pos(driver, 'linkCust_I_3', 1)
                                    # No se encontro
                                    send_click_pos(driver, 'linkCust_I_4', 0)
                                    # No se encontro
                                    send_click_pos(driver, 'linkCust_J_1', 0)

                                    # Stop If Stop Button is pressed
                                    if thread_stopped == True:
                                        break

                                    # F. Diagnosis and Plan - Diagnosis
                                    # ID: linkCust_F_1f
                                    # Name: Cust_F_1
                                    # Click in Yes
                                    res = send_click_pos(driver, 'linkCust_F_1', 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click F.1", True)

                                    # F. Diagnosis and Plan - Plan - a. Return to clinic: 
                                    # linkCust_F_a
                                    # Click in No
                                    res = send_click_pos(driver, 'linkCust_F_a', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click F.a", True)

                                    # F. Diagnosis and Plan - Plan - b. Minor fit to travel? 
                                    # linkCust_F_b
                                    # Click on Yes
                                    res = send_click_pos(driver, 'linkCust_F_b', 1)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click F.b", True)
                                    t.sleep(1)
                                    # F. Diagnosis and Plan - Plan - b1. Specify travel:
                                    try:
                                        driver.find_element(By.ID, 'linkCust_F_b1').clear()
                                        driver.find_element(By.ID, 'linkCust_F_b1').send_keys("The minor is medically cleared to travel if no known exposure to COVID has occurred and no other concerns requiring medical follow-up and/or specialty follow-up have been identified in subsequent visits.")
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click F.b1", True)
                                        pass

                                    # F. Diagnosis and Plan - Plan - c.Per program staff, discharge from ORR custody will be delayed?
                                    # linkCust_F_c
                                    # Click in No
                                    res = send_click_pos(driver, 'linkCust_F_c', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click F.c", True)

                                    # F. Diagnosis and Plan - Plan - d.Minor has/May have an ADA disability?
                                    # linkCust_F_d
                                    # Click in No
                                    res = send_click_pos(driver, 'linkCust_F_d', 0)
                                    if res:
                                        sendRequest(targent_name, "Error: Unable to click F.d", True)

                                    # F. Diagnosis and Plan - Plan - i.Age- appropriate anticipatory guidance discussed and/or handout given
                                    # linkCust_F_i
                                    # Clickbox
                                    # send_click_pos(driver, 'linkCust_F_i', 0)


                                    status_string = "Click button Save & Sign"
                                    print(status_string)
                                    # Click button Save & Sign
                                    try:
                                        wwwwww= driver.find_element_by_xpath('/html/body/table[5]/tbody/tr[2]/td/table/tbody/tr[8]/td/input[3]')
                                        ActionChains(driver).move_to_element(wwwwww).click(wwwwww).perform()
                                    except:
                                        sendRequest(targent_name, "Error: Unable to click Save & Sign", True)
                                        pass
                                    t.sleep(3)

                                    ## END - USER EDIT
                                    break
                                # else:
                                #     print("User not found")
                                #     # Click on search input field
                                #     search_select=driver.find_element_by_id('searchSelect')
                                #     ActionChains(driver).move_to_element(search_select).click(search_select).perform()

                                #     t.sleep(2)
                                #     try:
                                #         driver.find_element_by_xpath('//td[contains(text(),"All Facilities")]').click()
                                #     except:
                                #         pass
                                    
                                #     # Click on search input field
                                #     searchhhh_person=driver.find_element_by_id('searchField')
                                #     # Send Excel name
                                #     ActionChains(driver).move_to_element(searchhhh_person).click(searchhhh_person).send_keys(target_user_id, Keys.ENTER).perform()
                                    
                                #     # Stop If Stop Button is pressed
                                #     if thread_stopped == True:
                                #         break

                                #     # Select pop-up window "Global Resident Search -- All Residents"
                                #     t.sleep(5)
                                    
                                #     print(driver.window_handles)
                                #     try:
                                #         driver.switch_to.window(driver.window_handles[-1])
                                #         print("Select pop-up windo - Global Resident Search")
                                #     except:
                                #         pass

                                #     try:
                                #         driver.switch_to_window(driver.window_handles[-1])
                                #         print("Select pop-up windo - Global Resident Search2")
                                #     except:
                                #         pass

                                #     try:
                                #         driver.find_element_by_xpath('//a[contains(text(),"Current"]').click()
                                #         print("Click Current tab")
                                #     except:
                                #         pass

                                #     try:
                                #         driver.find_element_by_xpath("//a[@href='/admin/client/clientlist.jsp?ESOLview=Current&amp;ESOLglobalclientsearch=Y']").click();
                                #         print("Click Current tab2")
                                #     except:
                                #         pass

                                #     try:
                                #         driver.find_element_by_link_text("Current").click()
                                #         print("Click Current tab3")
                                #     except:
                                #         pass


                                t.sleep(1)
                                select_window(driver, 0)
                                print("END - USER EDIT")
                                ## END - USER EDIT
                                break
                            else:
                                print("User not found")
                                # Click on search input field
                                try:
                                    search_select=driver.find_element(By.ID, 'searchSelect')
                                    ActionChains(driver).move_to_element(search_select).click(search_select).perform()
                                except:
                                    pass

                                t.sleep(2)
                                try:
                                    driver.find_element(By.XPATH, '//td[contains(text(),"All Facilities")]').click()
                                except:
                                    pass
                                
                                # Click on search input field
                                # Send Excel name
                                res = send_text(driver, 'searchField', target_user_id)
                                if res:
                                    sendRequest(targent_name, "Error: Unable to click searchField", True)
                                res = send_enter(driver, 'searchField')
                                if res:
                                    sendRequest(targent_name, "Error: Unable to click searchField", True)
                                
                                # Stop If Stop Button is pressed
                                if thread_stopped == True:
                                    break

                                # Select pop-up window "Global Resident Search -- All Residents"
                                t.sleep(5)
                                
                                print(driver.window_handles)
                                try:
                                    driver.switch_to.window(driver.window_handles[-1])
                                    print("Select pop-up windo - Global Resident Search")
                                except:
                                    pass

                                try:
                                    driver.switch_to_window(driver.window_handles[-1])
                                    print("Select pop-up windo - Global Resident Search2")
                                except:
                                    pass

                                try:
                                    driver.find_element(By.XPATH, '//a[contains(text(),"Current"]').click()
                                    print("Click Current tab")
                                except:
                                    pass
                                try:
                                    driver.find_element(By.XPATH, "//a[@href='/admin/client/clientlist.jsp?ESOLview=Current&amp;ESOLglobalclientsearch=Y']").click();
                                    print("Click Current tab2")
                                except:
                                    pass

                                try:
                                    driver.find_element(By.LINK_TEXT, "Current").click()
                                    print("Click Current tab3")
                                except:
                                    pass

                                current_users=driver.find_element(By.CLASS_NAME, 'pccTableShowDivider').find_elements(By.TAG_NAME, 'tr')
                        except Exception as e:
                            print(e)
                            # sendRequest(targent_name, "Error: Unable to find user", True)
                            pass
                else:
                    print("NOT FOUND - " + target_user_id + "-" + target_user_a + "-" + targent_name)
                    res = write_file_data(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "," + str(target_user_id) + "," + str(target_user_a) + "," + targent_name + "\n")
                    if res:
                        sendRequest(targent_name, "Error: Unable to write not found error log", True)
                    
                    driver.close()
                    select_window(driver, 0)
                    t.sleep(1)
                    res = select_window(driver, 0)
                    if res:
                        sendRequest(targent_name, "Error: Could not switch to main window", True)
                    continue
            #except:
            except Exception as e:
                print(e)
                sendRequest(targent_name, "Error: Unable to select user", True)
                pass

        except Exception as e:
            print(e)
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.txt"), "a") as myfile:
                myfile.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " - " + str(e) + "\n")
            pass
    print("PROCESS COMPLETED")
    sendRequest("System", "Process Completed", False)
# Get
def get_covid_vaccine_by_name(imm_list):
    return imm_list[len(imm_list)-1]

def stop_automation_thread():
    global thread_stopped
    if thread_stopped == True:
        thread_stopped = False
    else:
        thread_stopped = True

def get_excel_gile():
    global patient_data_sheet

    # Open file explorer and select only xlsx file types
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("Excel file","*.xlsx")])

    patient_data_sheet = filename

class NewprojectApp:
    def __init__(self, master=None):
        global selected_sheet
        df = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx"), sheet_name=None)

        def start_automation_thread():
            # Set global variable to false
            global thread_stopped, user_name, password, selected_sheet
            thread_stopped = False
            # user_name = self.e1a.get()
            password = self.e2a.get()

            print('###############')
            # print(user_name)
            print(password)
            print('###############')

            # Create new thread target automation
            thread_x = Thread(target=main_loop, args=[])
            # Start Tread Activity
            thread_x.start()

        # build ui
        self.toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)

        # Title Label
        self.label1 = ttk.Label(self.toplevel1)
        self.label1.configure(background='#ff5454', font='{@Microsoft YaHei} 12 {}',
                              text='PCC')
        self.label1.pack(pady='10', side='top')

        self.frame2 = tk.Frame(self.toplevel1)


        # self.e1 = tk.Label(self.frame2)
        # self.e1.configure(background='#ffffff', text="User Name")
        # self.e1.pack()
        # self.e1a = tk.Entry(self.frame2)
        # self.e1a.pack()


        self.e2 = tk.Label(self.frame2)
        self.e2.configure(background='#ffffff', text="Password")
        self.e2.pack()
        self.e2a = tk.Entry(self.frame2)
        self.e2a.pack()

        # Open Chrome Button
        self.button5 = tk.Button(self.frame2)
        self.button5.configure( text='Open Chrome', command=open_chrome)
        self.button5.pack(ipadx='20', ipady='0', pady='5', side='top')

        # Sheets
        options = list(df.keys());
        clicked = tk.StringVar()
        clicked.set(options[0])
        selected_sheet = options[0];
        self.drop = tk.OptionMenu( self.frame2 , clicked , *options, command=selectSheet )
        self.drop.pack()

        # Start Button
        self.button6 = tk.Button(self.frame2)
        self.button6.configure(text='Start', command=start_automation_thread)
        self.button6.pack(ipadx='20', pady='5', side='top')
        # Stop Button
        self.button7 = tk.Button(self.frame2)
        self.button7.configure(text='Stop', command=stop_automation_thread)
        self.button7.pack(ipadx='20', pady='5', side='top')

        # Version Footer
        self.label2 = tk.Label(self.frame2)
        self.label2.configure(background='#ffffff', text="Version 2.3")
        self.label2.pack(side='top')
        self.frame2.configure(background='#ffffff', height='200', width='200')
        self.frame2.pack(side='top')

        # Window title bar
        self.toplevel1.configure(background='#ffffff', height='200', width='300')
        self.toplevel1.minsize(300, 200)
        self.toplevel1.overrideredirect('False')
        self.toplevel1.resizable(False, False)
        self.toplevel1.title('Pointclick Care Automator')

        # Main widget
        self.mainwindow = self.toplevel1

    def run(self):
        self.mainwindow.mainloop()

if __name__ == '__main__':
    app = NewprojectApp()
    app.run()
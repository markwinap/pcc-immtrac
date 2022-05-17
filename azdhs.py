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
get_e_timeout = 10
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

# Strings
specify_travel = "The minor is medically cleared to travel only if all covid quarantine clearance criteria have been met and no other concerns requiring medical follow up and/or specialty follow-up have been identified in subsequent visits."

def getNumber(txt):
  try:
    return int(txt)
  except Exception as e:
    return int(float(txt))

def dismiss_alert(d):
    try:
        WebDriverWait(d, 5).until(EC.alert_is_present())
        driver.switch_to.alert.dismiss()
    except:
        pass

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
    driver.get("https://asiis.azdhs.gov/")

def wait_button(d, el, t):
    try:
        el = WebDriverWait(d, get_e_timeout).until(EC.element_to_be_clickable((t, el)))
    except Exception as e:
        print(e)
        pass

def wait_window(d):
    try:
        WebDriverWait(d, get_e_timeout).until(EC.new_window_is_opened)
    except Exception as e:
        print(e)
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

def send_blind_text(d, data):
    try:
        d.send_keys(data)
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

def get_text_name(d, el):
    try:
        return d.find_element(By.NAME, el)
    except Exception as e:
        print(e)
        return ''
        pass

def get_parent(child):
    try:
        return child.find_element(By.XPATH, '..')
    except Exception as e:
        print(e)
        return ''
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

def send_click_by_value(d, el):
    try:
        d.find_element(By.XPATH, '//input[@value="' + el + '"]').click()
        return False
    except Exception as e:
        print(e)
        return True
        pass

def send_click_name(d, el):
    try:
        d.find_element(By.NAME, el).click()
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

def send_blind_enter(d):
    try:
        d.send_keys(Keys.ENTER)
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

def select_menu(d, id, value):
    try:
        s = Select(d.find_element(By.ID, id))
        s.select_by_value(value)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def select_menu_name_text(d, name, value):
    try:
        s = Select(d.find_element(By.NAME, name))
        s.select_by_visible_text(value)
        return False
    except Exception as e:
        print(e)
        return True
        pass

def overwrite_file():
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "users_not_found.txt"), "w") as myfile:
            myfile.write("userId,A,name\n")
    except Exception as e:
        print(e)
        pass

def write_file_data(data):
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "users_not_found.txt"), "a") as myfile:
            myfile.write(data)
    except Exception as e:
        print(e)
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

def getWordChars(str):
    try:
        return re.sub('[^\w]', '', str)
    except Exception as e:
        print(e)
        return ""

def getNumbers(str):
    try:
        return re.sub('[^0-9.]', '', str)
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
            "bot": "azdhs"
        }
        r = requests.post("https://2qpxr842pk.execute-api.us-east-1.amazonaws.com/Prod/post-sns-data", data=json.dumps(payload))
        return r
    except Exception as e:
        print(e)
        return ""

def main_loop():
    # read excel
    report = []
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
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            print("date and time =", dt_string, '>>>>logging in  ')	

            target_user_a = clean_text(str(getData(data,'A#')))
            target_user_id = clean_text(str(getData(data,'userId')))
            targent_name = getData(data,'First Name') + ' ' + getData(data,'Last Name')

            print("Processing - " + target_user_id + "-" + target_user_a + "-" + targent_name)
            select_window(driver, 0)
            t.sleep(1)
            res = select_window(driver, 0)
            if res:
                sendRequest(targent_name, "Error: Could not switch to main window", True)
                break


            driver.get("https://asiis.azdhs.gov/patient_search.jsp?direct=1")
            # Stop If Stop Button is pressed
            if thread_stopped == True:
                break

            # Patient Search
            wait_button(driver, "clearButton", By.ID)
            send_click(driver, "clearButton")

            if thread_stopped == True:
                break

            # Personal Info
            res = send_text(driver, "patientFirstName", getData(data,'First Name'))
            if res:
                sendRequest(targent_name, "Error: Unable to enter first name", True)
                break
            send_text(driver, "patientLastName", getData(data,'Last Name'))
            send_text(driver, "patientBirthDate", getData(data,'Last Name'))

            if thread_stopped == True:
                break

            # Birth Date
            b_day = pd.Timestamp(getData(data,'Birth Date'))
            res = send_text(driver, "patientBirthDate", b_day.strftime("%m/%d/%Y"))
            if res:
                sendRequest(targent_name, "Error: Unable to enter birth date", True)
                break
            send_enter(driver, "patientBirthDate")

            send_text(driver, "guardianFirstName", getData(data,'Program Name'))
            send_text(driver, "motherMaidenName", getData(data,'Program Name'))
            send_text(driver, "addressStreet", getData(data,'Program Name'))

            send_click(driver, "addressCountryCode_chzn")
            send_text(driver, "addressCountryCode_chzn_text", "United States")
            send_enter(driver, "addressCountryCode_chzn_text")
            t.sleep(1)
            send_text(driver,"addressZipCode", getData(data,'ZIP'))
            send_enter(driver, "addressZipCode")
            t.sleep(1)

            res = send_click(driver, "searchButton")
            if res:
                sendRequest(targent_name, "Error: Unable to click search button", True)
                break
            if thread_stopped == True:
                break
            
            # WIP
            updateUSer = clean_text(str(getData(data,'SIISID')))
            if (updateUSer == "") or (updateUSer == "0"):
                res = send_click(driver, "addPatientCheckbox")
                if res:
                    sendRequest(targent_name, "Error: Unable to click patient checkbox", True)
                    break

                wait_button(driver, "searchButton", By.ID)
                send_click(driver, "searchButton")
                if thread_stopped == True:
                    break

                wait_button(driver, "addPatientButton", By.ID)
                res = send_click(driver, "addPatientButton")
                if res:
                    sendRequest(targent_name, "Error: Unable to click add patient button", True)
                    break
                wait_button(driver, "saveButto", By.ID)

                if thread_stopped == True:
                    break

                #Gender FEMALE MALE  OTHER UNKNOWN
                res = select_menu_name_text(driver, "gender_code", getData(data,'Gender').strip().upper())
                if res:
                    sendRequest(targent_name, "Error: Unable to add gender information", True)
                    break

                res = select_menu_name_text(driver, "vfc_eligible_code", "State Program Eligibility")
                if res:
                    sendRequest(targent_name, "Error: Unable to select State Program Eligibility", True)
                    break
                res = send_click(driver, "aCommitButton")
                if res:
                    sendRequest(targent_name, "Error: Unable to click commit button", True)
                    break

                if thread_stopped == True:
                    break

                dismiss_alert(driver)
                send_click(driver, "gCommitButton")
                dismiss_alert(driver)
                # Save
                res = send_click(driver, "saveButton")
                if res:
                    sendRequest(targent_name, "Error: Unable to click save button", True)
                    break
            else:
                print("Update Patient " + updateUSer)
                #tableSearchResultsSortedThirdCol
                patientResults = driver.find_element(By.ID, "tableSearchResultsSortedThirdCol").find_elements(By.TAG_NAME, 'tr')
                
                print("Results " + str(len(patientResults)))
                for i in range(len(patientResults)):
                    row = patientResults[i + 1 ].find_elements(By.TAG_NAME, 'td')
                    print("Found ID " + clean_text(row[4].text))
                    patientId = getNumber(clean_text(row[4].text))
                    print("Patient ID " + str(patientId))
                    if patientId == getNumber(updateUSer):
                        patientResults[i + 1].click()
                        break

            wait_button(driver, "editHighRiskCategoriesButton", By.ID)

            if thread_stopped == True:
                break
            
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

                driver.get("https://asiis.azdhs.gov/vacc_view_add.jsp")
                wait_button(driver, '//input[@value="Add Administered"]', By.XPATH)

                for vacc in vaccines_list:

                    if thread_stopped == True:
                        break
                    select_vacc = get_vaccine_by_name(vacc, immunizations_list)
                    print("Add Vaccine Date")
                    print(select_vacc["Name"])

                    wordChars = getWordChars(select_vacc["azdhsId"]).lower()

                    print(wordChars)

                    print("Find Position")


                    vaccines_form = driver.find_element(By.ID, "vaccViewAdd")
                    vaccines_tables = vaccines_form.find_elements(By.CLASS_NAME, 'historylight')

                    row_number = 0
                    for i in range(len(vaccines_tables)):
                        if getWordChars(vaccines_tables[i].text).lower() == wordChars:
                            print("Found")
                            print(vaccines_tables[i].text)
                            print(vaccines_tables[i].get_attribute('id'))
                            row_number = getNumbers(vaccines_tables[i].get_attribute('id'))
                            break

                    # Date Administrated
                    v_day = pd.Timestamp(getData(data,'Date of visit'))
                    pos = 0
                    for d in range(5):
                        pos = d
                        try:
                            found = driver.find_elements(By.ID, "vacc" + row_number + "_" + str(d))
                            if len(found) > 0:
                                print("Found Space on " + "vacc" + row_number + "_" + str(pos))
                                break
                        except Exception as e:
                            print(e)
                            sendRequest(targent_name, str(e), True)
                            pass
                        
                    res = send_text(driver, "vacc" + row_number + "_" + str(pos), v_day.strftime("%m/%d/%Y"))
                    if res:
                        sendRequest(targent_name, "Error: "  + select_vacc["Name"] + " Unable to add vacine date", True)
                        break
                    
                    send_enter(driver, "vacc" + row_number + "_" + str(pos))

                if thread_stopped == True:
                    break
                # Save
                res = send_click_by_value(driver, "Add Administered")
                if res:
                    sendRequest(targent_name, "Error: unable to click add administered", True)
                    break

                #  VFC Eligibility Update
                t.sleep(2)
                wait_button(driver, "vfcEligibilityUpdateForm_0", By.ID)
                res = select_menu(driver, "vfcEligibilityUpdateForm_vfcCode", "34")
                if res:
                    sendRequest(targent_name, "Error: Unable to select VFC code", True)
                    break
                res = send_click(driver, "vfcEligibilityUpdateForm_0")
                if res:
                    sendRequest(targent_name, "Error: Unable to click VFC Eligibility", True)
                    break

                t.sleep(2)
                wait_button(driver, '//input[@value="Save"]', By.XPATH)
                for vacc in enumerate(vaccines_list):

                    if thread_stopped == True:
                        break                    
                    select_vacc = get_vaccine_by_name(vacc, immunizations_list)
                    print("Add Vaccine Manufacturer")
                    print(select_vacc["Name"] + "-" + clean_text(str(select_vacc["Menu val"])))

                    #find vaccine position
                    vaccine_pos = 1
                    for idy, vacc in enumerate(vaccines_list):
                        wordChars = getWordChars(select_vacc["azdhsId"]).lower()
                        child = get_text_name(driver, "vaccCode_" + str((idy + 1)))
                        parent = get_parent(child)
                        if getWordChars(parent.text).lower() == wordChars:
                            print("Found")
                            vaccine_pos = (idy + 1)
                            break

                    # Click on Manufacturer
                    print("manufacturerName_" + str(vaccine_pos))
                    res = send_click_name(driver, "manufacturerName_" + str(vaccine_pos))
                    if res:
                        sendRequest(targent_name, "Error: Unable to click vaccine manufacturer", True)
                        break
                    wait_window(driver)
                    t.sleep(2)
                    select_window(driver, -1)

                    lot_number = str(select_vacc["Lot#"])
                    print("LOT# "  + lot_number)

                    current_manufacture=driver.find_elements(By.TAG_NAME, 'tr')
                    manufacture_len=len(current_manufacture)
                    print(manufacture_len)
                    if manufacture_len > 4:
                        if thread_stopped == True:
                            break
                        print("Select Manufacture")
                        lot_options = []
                        not_found = True
                        pos = 0
                        for x in range(manufacture_len - 4):
                            row = current_manufacture[x + 2].find_elements(By.TAG_NAME, 'td')
                            #available = getNumber(row[6].text)
                            available = str(row[2].text)
                            print(available)
                            if lot_number == available:
                                pos = x
                                not_found = False
                                print("Found Lot Number")
                                break
                        #     print("available: " + str(available))
                        #     lot_options.append(available)

                        # largest_number = lot_options[0]
                        # for i in range(len(lot_options)):
                        #     if lot_options[i] > largest_number:
                        #         largest_number = lot_options[i]
                        #         pos = i
                        #print("largest_number: " + str(largest_number))

                        if not_found:
                            print("Lot # Not Found")
                            sendRequest(targent_name, "Error: Lot # Not Found " + lot_number, True)
                            send_click_by_value(driver, "Cancel")
                        else:
                            print("pos: " + str(pos))
                            manufacture = current_manufacture[pos + 2].find_elements(By.TAG_NAME, 'td')
                            select = manufacture[0].find_element(By.TAG_NAME, 'input')
                            ActionChains(driver).move_to_element(select).click(select).perform()
                    else:
                        print("NO ITEMS - " + select_vacc["Name"])
                        sendRequest(targent_name, "Error: No Vaccines Available - " + select_vacc["Name"], True)
                        send_click_by_value(driver, "Cancel")
                    # Go backa to main window
                    select_window(driver, 0)

                if thread_stopped == True:
                    break
                # Save Button
                print("Save")
                res = send_click_by_value(driver, "Save")
                if res:
                    sendRequest(targent_name, "Error: Unable to click Save Button", True)
                    break
                t.sleep(5)
                wait_button(driver, '//input[@value="Add Administered"]', By.XPATH)

            t.sleep(5)

        except Exception as e:
            print(e)
            sendRequest("System", str(e), True)
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
            # password = self.e2a.get()

            # print('###############')
            # print(user_name)
            # print(password)
            # print('###############')

            # Create new thread target automation
            thread_x = Thread(target=main_loop, args=[])
            # Start Tread Activity
            thread_x.start()

        # build ui
        self.toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)

        # Title Label
        self.label1 = ttk.Label(self.toplevel1)
        self.label1.configure(background='#fc9105', font='{@Microsoft YaHei} 12 {}',
                              text='AZDHS')
        self.label1.pack(pady='10', side='top')

        self.frame2 = tk.Frame(self.toplevel1)


        # self.e1 = tk.Label(self.frame2)
        # self.e1.configure(background='#ffffff', text="User Name")
        # self.e1.pack()
        # self.e1a = tk.Entry(self.frame2)
        # self.e1a.pack()


        # self.e2 = tk.Label(self.frame2)
        # self.e2.configure(background='#ffffff', text="Password")
        # self.e2.pack()
        # self.e2a = tk.Entry(self.frame2)
        # self.e2a.pack()

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
        self.label2.configure(background='#ffffff', text="Version 4.0.2")
        self.label2.pack(side='top')
        self.frame2.configure(background='#ffffff', height='200', width='200')
        self.frame2.pack(side='top')

        # Window title bar
        self.toplevel1.configure(background='#ffffff', height='200', width='300')
        self.toplevel1.minsize(300, 200)
        self.toplevel1.overrideredirect('False')
        self.toplevel1.resizable(False, False)
        self.toplevel1.title('AZDHS Automator')

        # Main widget
        self.mainwindow = self.toplevel1

    def run(self):
        self.mainwindow.mainloop()

if __name__ == '__main__':
    app = NewprojectApp()
    app.run()
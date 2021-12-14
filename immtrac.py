import time
import tkinter
import pickle
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import os
from datetime import datetime
import time as t
import pandas as pd
from selenium.webdriver.support.ui import Select
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox
from tkinter import filedialog as fd
from threading import Thread
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

# Global Variables
driver = None
patient_data_sheet = None
thread_stopped = False
get_e_timeout = 5
status_string = ""
options = None
selected_sheet = ""

# Strings
firs_name_el_id = "txtFirstName"
last_name_el_id = "txtLastName"
birth_date_el_id = "txtBirthDate"
gender_el_id = "optSexCode"
streen_address_el_id = "txtStreetAddress"
ethnicity_el_id = "lstEthnicity"
zip_el_id = "txtZip"
city_el_id = "txtCity"
country_el_id = "lstCounty"
date_administered_el_id = "vaccinationDate"
prescribed_by_el_id = "defaultAdministeredById"
save_button_el_id = "saveButton"
ethnicity="Hispanic or Latino"
race="Other Race"


# FUNCTIONS
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
    except Exception as e:
        print(e)
        pass

def send_text_name(d, el, data):
    try:
        d.find_element(By.NAME, el).clear()
        d.find_element(By.NAME, el).send_keys(data)
    except Exception as e:
        print(e)
        pass

def send_click_pos(d, el, pos):
    try:
        d.find_elements(By.ID, el)[pos].click()
    except Exception as e:
        print(e)
        pass

def send_click(d, el):
    try:
        d.find_element(By.ID, el).click()
    except Exception as e:
        print(e)
        pass

def send_click_name(d, el):
    try:
        d.find_elements(By.NAME, el).click()
    except Exception as e:
        print(e)
        pass

def send_click_link_text(d, el):
    try:
        d.find_element(By.PARTIAL_LINK_TEXT, el).click()
    except Exception as e:
        print(e)
        pass

def send_enter(d, el):
    try:
        d.find_element(By.ID, el).send_keys(Keys.ENTER)
    except Exception as e:
        print(e)
        pass

def click_link(d, el):
    try:
        d.find_element(By.XPATH, '//a[contains(text(),"' + el + '")]').click()
    except Exception as e:
        print(e)
        pass

def click_link_pos(d, el, pos):
    try:
        d.find_elements(By.XPATH, '//a[contains(text(),"' + el + '")]')[pos].click()
    except Exception as e:
        print(e)
        pass

def click_link_text(d, el):
    try:
        d.find_element(By.XPATH, '//a[text()="' + el + '"]').click()
    except Exception as e:
        print(e)
        pass

def click_link_href(d, el):
    try:
        d.find_element(By.XPATH, '//a[@href="' + el + '"]').click()
    except Exception as e:
        print(e)
        pass

def click_label_text(d, el):
    try:
        d.find_element(By.XPATH, '//label[text()="' + el + '"]').click()
    except Exception as e:
        print(e)
        pass

def click_button_value(d, el):
    try:
        d.find_element(By.XPATH, '//input[@value="' + el + '"]').click()
    except Exception as e:
        print(e)
        pass

def click_button_name(d, el):
    try:
        d.find_element(By.XPATH, '//input[@name="' + el + '"]').click()
    except Exception as e:
        print(e)
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
    except Exception as e:
        print(e)
        pass

def select_menu_index(d, name, value):
    try:
        s = Select(d.find_element(By.ID, name))
        s.select_by_index(value)
    except Exception as e:
        print(e)
        pass

def select_menu_name(d, name, value):
    try:
        s = Select(d.find_element(By.NAME, name))
        s.select_by_value(value)
    except Exception as e:
        print(e)
        pass

def select_menu_name_text(d, name, value):
    try:
        s = Select(d.find_element(By.NAME, name))
        s.select_by_visible_text(value)
    except Exception as e:
        print(e)
        pass

def select_menu_id_text(d, name, value):
    try:
        s = Select(d.find_element(By.ID, name))
        s.select_by_visible_text(value)
    except Exception as e:
        print(e)
        pass

def overwrite_file():
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "users_not_found.txt"), "w") as myfile:
            myfile.write("userId,A,name\n")
    except:
        pass

def write_file_data(data):
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "users_not_found.txt"), "a") as myfile:
            myfile.write(data)
    except:
        pass

def get_vaccine_by_name(vac_name, imm_list):
    for i in range(len(imm_list)):
        if imm_list[i]['Name'].lower() == vac_name.lower():
            return imm_list[i]

def write_pdf(client_id, a, name, last_name, cookie):
    try:
        # print('$$$$$$$$')
        # print(cookie)
        # print('$$$$$$$$')
        retry_strategy = Retry(
            total=3,
            status_forcelist=[429, 500, 502, 503, 504],
            method_whitelist=["HEAD", "GET", "OPTIONS"]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        http = requests.Session()
        http.mount("https://", adapter)
        http.mount("http://", adapter)

        url = "https://immtrac.dshs.texas.gov/TXPRD/auth/generateGroupPatientsReport.do?clientId=" + str(client_id) + "&reportTypeId=12"
        payload={}
        headers = {
        'Cookie': cookie
        }
        response = http.get(url, headers=headers)
        #response = requests.request("GET", url, headers=headers, data=payload)
        #'test.pdf'
        open("pdfs/"+ str(client_id) + "_" + str(a) + "_" + name.replace(" ", "").lower() + "_" + last_name.replace(" ", "").lower() + ".pdf", 'wb').write(response.content)
    except Exception as e:
        print(e)
        pass

def get_cookie_value(lis, name):
    try:
        c = list(filter(lambda c: c['name'] == name, lis))
        return c[0]['value']
    except:
        pass

def clean_text(txt):
    try:
        txt = txt.replace(".0", "")
        new_string = ''.join(char for char in txt if char.isalnum())
        return new_string
    except:
        pass

# END FUNCTIONS


def update_status(msg):
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print(current_time + " - " + msg)

def open_chrome():
    global driver, options
    tkinter.messagebox.showinfo("Information", "Please log into your account using next opening browser. Then click on 'Start' button to start the automation.")
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
    # Open website
    driver.get("https://immtrac.dshs.texas.gov/TXPRD/portalHeader.do")

def main_loop():
    # read excel
    global driver,patient_data_sheet,status_string, options, selected_sheet

    print("SELECTED SHEET " + selected_sheet)

    # Read Immunizations.xlsx Excel file
    df = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx"), sheet_name=selected_sheet)
    patient_list = []
    for index, row in df.iterrows():
        patient_list.append(row)

    # Read Immunizations.xlsx Excel file
    df = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Immunizations.xlsx"), sheet_name='Sheet1')
    imm_list = []
    for index, row in df.iterrows():
        imm_list.append(row)

    # Loop through patient list
    for i in range(len(patient_list)):
        if thread_stopped == True:
            break
        time.sleep(3)
        try:
            # driver.execute_script("openLink('search_ui.findClientForNew');")
            # Click on dissmiss
            try:
                WebDriverWait(driver, get_e_timeout).until(EC.alert_is_present())
                driver.switch_to.alert.dismiss()
            except:
                pass
            if str(patient_list[i]['ImmTrac2Id']) != '' and len(str(patient_list[i]['ImmTrac2Id'])) > 1 and str(patient_list[i]['ImmTrac2Id']) != 'nan':
                print(patient_list[i]['ImmTrac2Id'])
                print(str(patient_list[i]['ImmTrac2Id']))
                print(clean_text(str(patient_list[i]['ImmTrac2Id'])))
                
                print('Search Client By Id')
                click_link_pos(driver, 'manage client', 1)
                t.sleep(2)
                try:
                    driver.execute_script("toggleInfoBlock('quickDiv', 'quickArrow', quickDiv);");
                except:
                    pass

                t.sleep(0.5)

                # Enter ImmTrac2Id "txtWirID"
                send_text(driver, 'txtWirID', clean_text(str(patient_list[i]['ImmTrac2Id'])))
                send_click_pos(driver, 'cmdFind', 0)
                t.sleep(2)
                status_string = "Wait till Immunizations is available"
                print(status_string)
                try:
                    el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Immunizations']")))
                except:
                    print("NOT FOUND - " + str(patient_list[i]['ImmTrac2Id']) + "-" + clean_text(str(patient_list[i]['First Name'])))
                    write_file_data(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "," + clean_text(str(patient_list[i]['ImmTrac2Id'])) + "," + str(patient_list[i]['First Name']) + "\n")
                    continue
                    pass
                    #No clients were found for the requested search criteria.
                click_button_value(driver, 'Immunizations')
                # Stop If Stop Button is pressed
                if thread_stopped == True:
                    break

                # Wait for 2 Second
                t.sleep(2)

            else:
                print('Enter New Client')
                # Click on Button 'enter new client'
                click_link(driver, 'enter new client')
                click_link(driver, '&nbsp;enter new client')

                # Wait until an element with id='txtFirstName' is available with a timeout opf 10 seconds
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.ID, firs_name_el_id)))
                except:
                    pass

                # Wait 1 Second (While Page Loads)
                t.sleep(1)
                status_string = "Client Search - " + patient_list[i]['First Name'] + " " + patient_list[i]['Last Name']
                print(status_string)
                ## Step 1 -  Client Search
                # First Name
                send_text(driver, firs_name_el_id, patient_list[i]['First Name'])
                # Last Name
                send_text(driver, last_name_el_id, patient_list[i]['Last Name'])
                # Birth Date
                b_day = pd.Timestamp(patient_list[i]['Birth Date'])
                send_text(driver, birth_date_el_id, b_day.strftime("%m/%d/%Y"))
                # Gender
                select_menu_id_text(driver, gender_el_id, patient_list[i]['Gender'])
                # Address
                send_text(driver, streen_address_el_id, patient_list[i]['Street Address'])
                # Click on Find button
                send_click(driver, 'cmdFind')

                # Wait for 1 Second
                t.sleep(1)
                status_string = "Do you have one of the following signed consent forms for your client?"
                print(status_string)
                ## Step 2 - Confirm Add New Client
                #  Wait until an element with name "optConsentR" - Do you have one of the following signed consent forms for your client? - timeout 12 seconds
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='optConsentR']")))
                except:
                    pass
                status_string = "Click Add a client and Submit"
                print(status_string)
                # Click on Yes - Add a client
                click_label_text(driver, 'Add a client')
                # Click On Submit Button
                click_button_value(driver, 'Submit')
                # Wait for 2 Second
                t.sleep(2)
                status_string = "Wait until an label element with name chkRace is available"
                print(status_string)
                ## Step 3 - Client Information
                #  Wait until an label element with name "chkRace" is available - timeput 12 seconds
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//label[@for='chkRace']")))
                except:
                    pass

                # Stop If Stop Button is pressed
                if thread_stopped == True:
                    break

                status_string = "Select Race"
                print(status_string)
                # Select Race base if text contains
                if "indian" in race.lower():
                    click_button_name(driver, 'indian')
                if "asian" in race.lower():
                    click_button_name(driver, 'asian')
                if "hawaiian" in race.lower():
                    click_button_name(driver, 'hawaiian')
                if "black" in race.lower():
                    click_button_name(driver, 'black')
                if "white" in race.lower():
                    click_button_name(driver, 'white')
                if "other" in race.lower():
                    click_button_name(driver, 'other')
                if "recipient" in race.lower():
                    click_button_name(driver, 'recipientRefused')

                # Stop If Stop Button is pressed
                if thread_stopped == True:
                    break

                # Select Ethnicity
                ## Hispanic or Latino
                ## Not Hispanic or Latino
                ## Recipient Refused
                status_string = "Select Ethnicity"
                print(status_string)
                select_menu_id_text(driver, ethnicity_el_id, ethnicity)
                # ZIP Code
                send_text(driver, zip_el_id, patient_list[i]['ZIP'])
                # City
                send_text(driver, city_el_id, patient_list[i]['City'])
                # Country
                select_menu_id_text(driver, country_el_id, patient_list[i]['County'])

                # Stop If Stop Button is pressed
                if thread_stopped == True:
                    break

                status_string = "Click on Continue Add"
                print(status_string)
                # Click On Button With value "Continue Add"
                click_button_value(driver, 'Continue Add')
                # Wait for 2 Second
                t.sleep(2)
                ## Step 4 - Consent Affirmation
                #  Wait until button with value "Continue" is available - timeput 12 seconds
                status_string = "Wait till Continue button is available"
                print(status_string)
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Continue']")))
                except:
                    pass
                # Click on Continue button
                click_button_value(driver, 'Continue')
                # Wait for 6 Second
                t.sleep(6)

                status_string = "Wait till Create New Client button is available"
                print(status_string)
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Create New Client']")))
                    click_button_value(driver, 'Create New Client')
                except:
                    pass

                ################################
                #  Wait until button with value "Affirm" is available - timeput 12 seconds
                status_string = "Wait till Affirm button is available"
                print(status_string)
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Affirm']")))
                    # Click on "Affirm" button
                    click_button_value(driver, 'Affirm')
                except:
                    pass

                status_string = "Wait till Create New Client button is available"
                print(status_string)
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Create New Client']")))
                    click_button_value(driver, 'Create New Client')
                except:
                    pass
                try:
                    WebDriverWait(driver, get_e_timeout).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                except:
                    pass

                status_string = "Wait till Affirm button is available"
                print(status_string)
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Affirm']")))
                    # Click on "Affirm" button
                    click_button_value(driver, 'Affirm')
                except:
                    pass

                # Wait for 6 Second
                t.sleep(6)
                status_string = "Wait till Go To client button is available"
                print(status_string)
                #  Wait until button with value "Go To client" is available - timeput 12 seconds (NEED TO CHECK)

                try:
                    # Click on Button 'enter new client'
                    driver.find_element_by_xpath("//input[contains(text(),'Go To client')]").click()
                except:
                    driver.find_element_by_xpath("//button[@onclick='gotoClient(); return false;']").click()
                    pass

                # Wait for 2 Second
                t.sleep(6)
                status_string = "Wait till Immunizations is available"
                print(status_string)
                #  Wait until button with value "Immunizations" is available - timeput 12 seconds
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Immunizations']")))
                except:
                    driver.find_element_by_xpath("//*[contains(text(), 'Go To client')]").click()
                    t.sleep(6)
                    pass
                # Click on "Immunizations" button
                click_button_value(driver, 'Immunizations')

                # Stop If Stop Button is pressed
                if thread_stopped == True:
                    break

                # Wait for 2 Second
                t.sleep(2)

            ## Step 5 - Immunizations
            #  Wait until button with value "Add New Imms" is available - timeput 12 seconds
            status_string = "Wait till Add New Imms is available"
            print(status_string)
            try:
                el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Add New Imms']")))
            except:
                pass
            # Click on "Add New Imms" button
            click_button_value(driver, 'Add New Imms')
            #  Wait until field with id "vaccinationDate" is available - timeput 12 seconds
            try:
                el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.ID, date_administered_el_id)))
            except:
                pass

            print('Populate Array')
            # Update Vactinations Array
            vaccine_list_to_add = []
            if patient_list[i]['Hepatitis A'] == "Yes":
                vaccine_list_to_add.append("Hepatitis A")
            if patient_list[i]['Hepatitis B'] == "Yes":
                vaccine_list_to_add.append("Hepatitis B")
            if patient_list[i]['HPV'] == "Yes":
                vaccine_list_to_add.append("HPV")
            if patient_list[i]['Influenza'] == "Yes":
                vaccine_list_to_add.append("Influenza")
            if patient_list[i]['Meningicoccal'] == "Yes":
                vaccine_list_to_add.append("Meningicoccal")
            # if patient_list[i]['Td/Tdap'] == "Yes":
            #     vaccine_list_to_add.append("Td/Tdap")
            if patient_list[i]['Td'] == "Yes":
                vaccine_list_to_add.append("Td")
            if patient_list[i]['Tdap'] == "Yes":
                vaccine_list_to_add.append("Tdap")
            if patient_list[i]['Varicella'] == "Yes":
                vaccine_list_to_add.append("Varicella")
            if patient_list[i]['MMR'] == "Yes":
                vaccine_list_to_add.append("MMR")
            if patient_list[i]['IPV'] == "Yes":
                vaccine_list_to_add.append("IPV")
            if patient_list[i]['SARS-COV-2'] == "Yes":
                vaccine_list_to_add.append("SARS-COV-2")
            if patient_list[i]['SARS-COV-2 < 12'] == "Yes":
                vaccine_list_to_add.append("SARS-COV-2 < 12")
            print(vaccine_list_to_add)

            # Date Administered
            date_admin_1 = pd.Timestamp(patient_list[i]['Date of visit'])
            send_text(driver, date_administered_el_id, date_admin_1.strftime("%m/%d/%Y"))

            # Prescribed By
            select_menu_index(driver, prescribed_by_el_id, 1)
            # Stop If Stop Button is pressed
            if thread_stopped == True:
                break

            # Get min numbe of elements to add
            range_x = 0
            if len(vaccine_list_to_add) <= 6:
                range_x = len(vaccine_list_to_add)
            else:
                range_x = 6

            status_string = "Fill vaccines"
            print(status_string)
            # Loop first 6 elements
            for xx in range(range_x):
                if thread_stopped == True:
                    break
                try:
                    # Immunization - Select "Immunization name" from Excel
                    select_menu_name_text(driver, "vaccineGroupId[" + str(xx) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['NameImmtrac'])
                    t.sleep(0.5)
                    # Trade Name
                    select_menu_name_text(driver, "tradeNameId[" + str(xx) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['Trade Name'])
                    t.sleep(0.5)
                    # Lot # 
                    select_3 = driver.find_element_by_xpath("//input[@name='historicalLotNumber[" + str(xx) + "]']")
                    select_3.send_keys(get_vaccine_by_name(vaccine_list_to_add[xx], imm_list)['Lot#'])
                    # Vaccine Eligibility - Error Excel missing options
                    select_menu_name_text(driver, "vaccineEligibilityCode[" + str(xx) + "]", "V03-No Insurance")
                    # Prescribed By -  Default Juan Garcia
                    select_5 = Select(driver.find_element_by_xpath("//select[@name='administeredById[" + str(xx) + "]']"))
                    select_5.select_by_index(1)
                    # Manufacturer
                    select_menu_name_text(driver,  "manufacturerId[" + str(xx) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['ManufacturerImmtrac'])
                    # Body Site
                    select_menu_name_text(driver,  "bodySiteCode[" + str(xx) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['Zone'].upper() + ' ARM')
                    # Route
                    select_menu_name_text(driver,  "adminRouteCode[" + str(xx) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['Route'])

                except Exception as e:
                    # log error
                    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.txt"), "a") as myfile:
                        myfile.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " - " + vaccine_list_to_add[
                            xx] + " - " + str(e) + "\n")
                    pass

            # Stop If Stop Button is pressed
            if thread_stopped == True:
                break
            
            # Click On Save Button
            send_click(driver, save_button_el_id)

            # If patient Covid vaccine is marked as Yes or vaccines to add is > 6
            if len(vaccine_list_to_add) > 6:
                ## Step 6 - Add more vaccine 
                #  Wait until button with value "Add New Imms" is available - timeput 12 seconds
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@value='Add New Imms']")))
                except:
                    pass
                # Click On Add New Imms button
                click_button_value(driver, 'Add New Imms')                
                #  Wait until field with id "vaccinationDate" is available - timeput 12 seconds
                try:
                    el = WebDriverWait(driver, get_e_timeout).until(EC.element_to_be_clickable((By.ID, date_administered_el_id)))
                except:
                    pass

                # # Date Administered
                #driver.find_element_by_id(date_administered_el_id).send_keys(datetime.now().strftime("%m/%d/%Y"))
                date_admin_2 = pd.Timestamp(patient_list[i]['Date of visit'])
                send_text(driver, date_administered_el_id, date_admin_2.strftime("%m/%d/%Y"))
                # Prescribed By
                select_menu_index(driver, 'defaultAdministeredById', 1)

                # Wait for 2 Second
                t.sleep(2)
                range_x = len(vaccine_list_to_add)

                if thread_stopped == True:
                    break

                for xx in range(6, range_x):
                    if thread_stopped == True:
                        break
                    try:
                        # Immunization - Select "Immunization name" from Excel
                        select_menu_name_text(driver, "vaccineGroupId[" + str(xx - 6) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['NameImmtrac'])
                        t.sleep(0.5)
                        # Trade Name
                        select_menu_name_text(driver, "tradeNameId[" + str(xx - 6) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['Trade Name'])
                        t.sleep(0.5)
                        # Lot # 
                        select_3 = driver.find_element_by_xpath("//input[@name='historicalLotNumber[" + str(xx - 6) + "]']")
                        select_3.send_keys(get_vaccine_by_name(vaccine_list_to_add[xx], imm_list)['Lot#'])
                        # Vaccine Eligibility - Error Excel missing options
                        select_menu_name_text(driver, "vaccineEligibilityCode[" + str(xx - 6) + "]", "V03-No Insurance")
                        # Prescribed By -  Default Juan Garcia
                        select_5 = Select(driver.find_element_by_xpath("//select[@name='administeredById[" + str(xx - 6) + "]']"))
                        select_5.select_by_index(1)
                        # Manufacturer
                        select_menu_name_text(driver,  "manufacturerId[" + str(xx - 6) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['ManufacturerImmtrac'])
                        # Body Site
                        select_menu_name_text(driver,  "bodySiteCode[" + str(xx - 6) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['Zone'].upper() + ' ARM')
                        # Route
                        select_menu_name_text(driver,  "adminRouteCode[" + str(xx - 6) + "]", get_vaccine_by_name(vaccine_list_to_add[xx],imm_list)['Route'])
                    except Exception as e:
                        # log error
                        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.txt"), "a") as myfile:
                            myfile.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " - " + vaccine_list_to_add[xx] + " - " + str(e) + "\n")
                        pass
                if thread_stopped == True:
                    break
                try:
                    # Click On Save Button
                    send_click(driver,save_button_el_id)
                    status_string = "Click on save button"
                    print(status_string)
                except:
                    pass
                if thread_stopped == True:
                    break
            
            # Click on dissmiss
            try:
                WebDriverWait(driver, get_e_timeout).until(EC.alert_is_present())
                driver.switch_to.alert.dismiss()
            except Exception as e:
                print(e)
                pass
            print("Save PDF File")
            cookies = driver.get_cookies()
            client_id = driver.find_element(By.NAME, 'txtClientId').get_attribute('value')
            print(str(patient_list[i]['A#']) + "_" + patient_list[i]['First Name'])
            write_pdf(client_id, patient_list[i]['A#'], patient_list[i]['First Name'], patient_list[i]['Last Name'], 'ROUTEID=' + get_cookie_value(cookies, 'ROUTEID') + '; dtCookie=' + get_cookie_value(cookies, 'dtCookie') + '; rxVisitor=' + get_cookie_value(cookies, 'rxVisitor') + '; iisjsessionid=' + get_cookie_value(cookies, 'iisjsessionid') + '; dtSa=' + get_cookie_value(cookies, 'dtSa') + '; dtLatC=' + get_cookie_value(cookies, 'dtLatC') + '; rxvt=' + get_cookie_value(cookies, 'rxvt') + '; dtPC=' + get_cookie_value(cookies, 'rxvt'))

            if len(driver.window_handles) > 1:
                select_window(driver, -1)
                driver.close()
                select_window(driver, 0)


        except Exception as e:
            # log error
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.txt"), "a") as myfile:
                myfile.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " - " + "Main profile creation" + " - " + str(e) + "\n")
            pass
        time.sleep(4)
    print("COMPLETED")
# 
def get_vaccine_by_name(vac_name, imm_list):
    for i in range(len(imm_list)):
        if imm_list[i]['Option'].lower() == vac_name.lower():
            return imm_list[i]

# Get
def get_covid_vaccine_by_name(imm_list):
    return imm_list[len(imm_list)-1]

def start_automation_thread():
    # Set global variable to false
    global thread_stopped
    thread_stopped = False
    # Create new thread target automation
    thread_x = Thread(target=main_loop, args=[])
    # Start Tread Activity
    thread_x.start()

def stop_automation_thread():
    global thread_stopped
    if thread_stopped == True:
        thread_stopped = False
    else:
        thread_stopped = True

def selectSheet(sheet):
    global selected_sheet
    selected_sheet = sheet
    print(selected_sheet)

def get_excel_gile():
    global patient_data_sheet

    # Open file explorer and select only xlsx file types
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("Excel file","*.xlsx")])

    patient_data_sheet = filename

# Create UI form
class NewprojectApp:
    def __init__(self, master=None):
        global status_string, selected_sheet

        df = pd.read_excel(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx"), sheet_name=None)
        # build ui
        self.toplevel1 = tk.Tk() if master is None else tk.Toplevel(master)
        # Title Label
        self.label1 = ttk.Label(self.toplevel1)
        self.label1.configure(background='#336699', font='{@Microsoft YaHei} 12 {}',
                              text='Immtrac Automator', foreground='#ffffff')
        self.label1.pack(pady='10', side='top')

        self.frame2 = tk.Frame(self.toplevel1)


        # # Select Excel Sheet Button
        # self.open_button = ttk.Button(
        #     self.frame2,
        #     text='Select Excel sheet (patient details)',
        #     command=get_excel_gile
        # )
        # self.open_button.pack(ipadx='20', ipady='0', pady='5', side='top')

        # Open Chrome Button
        self.button5 = tk.Button(self.frame2)
        self.button5.configure( text='Open Chrome', command=open_chrome)
        self.button5.pack(ipadx='20', ipady='0', pady='5', side='top')

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
        self.label2.configure(background='#ffffff', text="Version 1.7")
        self.label2.pack(side='top')
        self.frame2.configure(background='#ffffff', height='200', width='200')
        self.frame2.pack(side='top')

        # Window title bar
        self.toplevel1.configure(background='#ffffff', height='200', width='300')
        self.toplevel1.minsize(300, 200)
        self.toplevel1.overrideredirect('False')
        self.toplevel1.resizable(False, False)
        self.toplevel1.title('Immtrac Automators')

        # Main widget
        self.mainwindow = self.toplevel1

    def run(self):
        self.mainwindow.mainloop()

if __name__ == '__main__':
    app = NewprojectApp()
    app.run()
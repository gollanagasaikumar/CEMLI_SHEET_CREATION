import os
import sys
import time
import datetime
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options  
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException


CREDENTIALS_LIST = ["nagasaikumar.golla@geappliances.com","515120537","NAsa@321ku"]
now = datetime.now()
date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
CEMLI_NAME = os.getenv("CEMLI_SHEET")
#os.getenv("CEMLI_SHEET")
#driver = webdriver.Chrome(r"C:\Users\nagasaikumar.golla\Desktop\CEMLI_SHEET_AUTOMATION_SCRIPT\chromedriver.exe")
chrome_options = Options()  
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1366,768")
driver = webdriver.Chrome(r"C:\Users\nagasaikumar.golla\Desktop\CEMLI_SHEET_AUTOMATION_SCRIPT\chromedriver.exe",options=chrome_options)
time.sleep(3)
driver.maximize_window()
print("***************** CEMLI Creation Started for " + CEMLI_NAME + "*****************  \n ")
sys.stdout.flush()
driver.get('https://geappliances.sharepoint.com/sites/erpdevops/SitePages/SOA.aspx')
try:
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": CEMLI Sheet Creation Started \n ")	
    sys.stdout.flush()
    inputElement = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "i0116")))
    inputElement.send_keys(CREDENTIALS_LIST[0])
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Enter GE Appliances Mail ID \n ")	
    sys.stdout.flush()
    submit_button = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
    submit_button.click()
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Mail ID Submit Button \n ")	
    sys.stdout.flush()
    SSOID = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.NAME, "username")))
    SSOID.send_keys(CREDENTIALS_LIST[1])
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Enter SSO ID \n ")	
    sys.stdout.flush()
    SSOPWD = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.NAME, "password")))
    SSOPWD.send_keys(CREDENTIALS_LIST[2])
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Enter Password  \n ")	
    sys.stdout.flush()
    SSOSUBMIT = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.NAME, "submit")))
    SSOSUBMIT.click()
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Login Button  \n ")	
    sys.stdout.flush()
    Remember = WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
    Remember.click()
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Login Authentication Completed  \n ")
    sys.stdout.flush()
    AddNewPage = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "idHomePageNewWikiPage")))
    AddNewPage.click()
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Add New Cemli Page  \n ")	
    sys.stdout.flush()
    EnterCemliName = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "ctl00_PlaceHolderMain_nameInput")))
    EnterCemliName.send_keys(CEMLI_NAME)
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Enter CEMLI Name  \n ")	
    sys.stdout.flush()
    CemliCreate = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "ctl00_PlaceHolderMain_createButton")))
    CemliCreate.click()
    print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Create Cemli Page \n ")	
    sys.stdout.flush()
    if driver.find_elements_by_xpath("//span[contains(@class,'ms-error') and .//text()='The specified name is already in use.']"):
        print("Entered Cemli Page Already Exists \n https://geappliances.sharepoint.com/sites/erpdevops/SitePages/"+CEMLI_NAME+".aspx  \n ")
        sys.stdout.flush()
    else:
        #driver.save_screenshot('./foto.png')
        #EditSourceClick1 = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.Markup-PopupAnchor-Large")))
        #EditSourceClick1.click()
        driver.save_screenshot('./foto1.png')
        EditSourceClick = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.Markup.Html.Menu.Html.EditSource-Large")))
        EditSourceClick.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Edit Source \n ")	
        sys.stdout.flush()
        driver.save_screenshot('./foto2.png')
        HTMLCode = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "PropertyEditor")))
        HTMLCode.clear()
        x = '<table width="1313" cellspacing="0" height="954" class="ms-rteTable-0"> <tbody> <tr class="ms-rteTableEvenRow-0"> <td class="ms-rteTableEvenCol-0" style="width: 50%;"> <table width="100%" cellspacing="0" class="ms-rteTable-default"> <tbody> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default" colspan="2" style="width: 50%;"> <strong class="ms-rteFontSize-3">INFORMATION</strong></td> </tr> <tr class="ms-rteTableOddRow-default"> <td valign="top" class="ms-rteTableEvenCol-default" style="width: 13%;"> <strong>JIRA Ticket</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> <br/> </td> </tr> <tr class="ms-rteTableEvenRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Name</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> <br/> </td> </tr> <tr class="ms-rteTableEvenRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Description</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> <br/> </td> </tr> <tr class="ms-rteTableOddRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Source</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> Source CI: <br/>Database Table: <br/>Directory: <br/>File: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Integration</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default">Integration CI: <br/>Objects: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Target</strong></td> <td class="ms-rteTableOddCol-default">Target CI: <br/>API: <br/>Database Table: <br/>Directory: <br/>File: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Dependencies</strong></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>External</strong></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Contacts</strong></td> <td class="ms-rteTableOddCol-default">Source: <br/>Integration: <br/>Target: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Reprocessing</strong></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Regression Testing</strong><br/></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Support Transition Status</strong><br/></td> <td class="ms-rteTableOddCol-default">MD2060: <br/>Code Walk Through: <br/>CEMLI Page: <br/>Final Sign Off: <br/> </td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Sample Payload</strong><br/></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>non-PROD Details</strong><br/></td> <td class="ms-rteTableOddCol-default">dev: <br/>tst: <br/>qa/stg: <br/></td> </tr> </tbody> </table> </td> <td class="ms-rteTableOddCol-0" style="width: 50%;"> <table width="100%" cellspacing="0" class="ms-rteTable-default"> <tbody> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default" colspan="2" style="width: 50%;"> <strong><span><span><strong class="ms-rteFontSize-3">FAILURE MODES<br/></strong></span></span></strong></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default" style="width: 5%;">#1</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#2</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#3</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#4</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#5</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#6</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#7</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#8</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#9</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableFooterRow-default"> <td class="ms-rteTableFooterEvenCol-default" rowspan="1">#10</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> </tbody> </table> </td> </tr> </tbody></table><p> <br/></p>'
        HTMLCode.send_keys(x)
        time.sleep(3)
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Enter HTML Code \n ")	
        sys.stdout.flush()
        HTMLCodeSubmit = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "sourcedialog_okbutton")))
        HTMLCodeSubmit.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on HTML Code Submit \n ")	
        sys.stdout.flush()
        Checkin = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.EditAndCheckout.Checkout-SelectedItem")))
        Checkin.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Check-In \n ")	
        sys.stdout.flush()
        CheckinContinue = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "statechangedialog_okbutton")))
        CheckinContinue.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Check-In Continue PopUp \n ")	
        sys.stdout.flush()
        ContentType = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select an option']")))
        ContentType.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Content Type \n ")	
        sys.stdout.flush()
        ContentTypeSelection = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='CEMLIs']")))
        ContentTypeSelection.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Selecting Content Type \n ")	
        sys.stdout.flush()
        Track = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select options']")))
        Track.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Track \n ")	
        sys.stdout.flush()
        TrackSelection = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='Shared']")))
        TrackSelection.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Selecting Track Type \n ")	
        sys.stdout.flush()
        Module = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select options']")))
        Module.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Module \n ")	
        sys.stdout.flush()
        ModuleSelection = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='Sourcing']")))
        ModuleSelection.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Selecting Module Type \n ")	
        sys.stdout.flush()
        JustClick = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//label[contains(@class,'ms-Label ReactFieldEditor-fieldTitle') and .//text()='Description']")))
        JustClick.click()
        Description = driver.find_elements_by_xpath("//input[@placeholder='Enter value here']")[1].send_keys(CEMLI_NAME)
        time.sleep(2)
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Enter Description Details \n ")	
        sys.stdout.flush()
        Platform = driver.find_elements_by_xpath("//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select options']")[1].click()
        time.sleep(2)
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Platform \n ")	
        sys.stdout.flush()
        PlatformSelection = driver.find_elements_by_xpath("//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='SOA']")[0].click()
        time.sleep(3)
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Select Platform Type \n ")	
        sys.stdout.flush()
        DropdownClose = driver.find_elements_by_xpath("//span[contains(@class,'ms-Dropdown-caretDown') and .//text()='']")[5].click()
        time.sleep(2)
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Completed Filling Data \n ")	
        sys.stdout.flush()
        SaveCEMLI = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Button-label') and .//text()='Save']")))
        SaveCEMLI.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Save / Create Cemli \n ")	
        sys.stdout.flush()
        Edit = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-promotedActionButton-text') and .//text()='Edit']")))
        Edit.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Edit Cemli for Final Check-Out \n ")	
        sys.stdout.flush()
        Checkout = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.EditAndCheckout.Checkout-SelectedItem")))
        Checkout.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Check-Out \n ")	
        sys.stdout.flush()
        Finalcheckout = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "statechangedialog_okbutton")))
        Finalcheckout.click()
        print(now.strftime("%m/%d/%Y, %H:%M:%S") + ": Click on Check-Out Confirmation \n ")	
        sys.stdout.flush()
        print (now.strftime("%m/%d/%Y, %H:%M:%S") + ": CEMLI Sheet has been created succesfully please find the below CEMLI Sheet Link \n ")
        sys.stdout.flush()
        print("https://geappliances.sharepoint.com/sites/erpdevops/SitePages/"+CEMLI_NAME+".aspx")
        sys.stdout.flush()
except NoSuchElementException:
    print (now.strftime("%m/%d/%Y, %H:%M:%S") + ": Error Message :- Element Not Found and Execution got Failed \n Reason :- Element we are trying to access is not found or Slow Network Connection Time Out")
    sys.stdout.flush()
except StaleElementReferenceException:
    print (now.strftime("%m/%d/%Y, %H:%M:%S") + ": Error Message :- Stale Element reference \n Reason :- Stale Element means an old element or no longer available element")
    sys.stdout.flush()
except:
    print (now.strftime("%m/%d/%Y, %H:%M:%S") + ": Error Message :- Some Thing Went Wrong or Flow Terminated abruptly  \n")
    sys.stdout.flush()
finally:
    print ("***************** CEMLI Creation Completed *****************")
    sys.stdout.flush()
		

			
			

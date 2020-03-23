import os
import sys
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options  
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException



CREDENTIALS_LIST = ["nagasaikumar.golla@geappliances.com","515120537","NAsa@321ku"]
CEMLI_NAME = os.getenv("CEMLI_SHEET")
#driver = webdriver.Chrome(r"C:\Users\nagasaikumar.golla\Desktop\CEMLI_SHEET_AUTOMATION_SCRIPT\chromedriver.exe")
chrome_options = Options()  
chrome_options.add_argument("--headless")
driver = webdriver.Chrome(r"C:\Users\nagasaikumar.golla\Desktop\CEMLI_SHEET_AUTOMATION_SCRIPT\chromedriver.exe",options=chrome_options)
time.sleep(3)
driver.maximize_window()
print("***************** CEMLI Creation Started for " + CEMLI_NAME + "*****************")
sys.stdout.flush()
driver.get('https://geappliances.sharepoint.com/sites/erpdevops/SitePages/SOA.aspx')
try:
    inputElement = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "i0116")))
    inputElement.send_keys(CREDENTIALS_LIST[0])
    submit_button = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
    submit_button.click()
    SSOID = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.NAME, "username")))
    SSOID.send_keys(CREDENTIALS_LIST[1])
    SSOPWD = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.NAME, "password")))
    SSOPWD.send_keys(CREDENTIALS_LIST[2])
    SSOSUBMIT = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.NAME, "submit")))
    SSOSUBMIT.click()
    print("Login Authentication Completed")
    sys.stdout.flush()
    Remember = WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
    Remember.click()
    AddNewPage = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "idHomePageNewWikiPage")))
    AddNewPage.click()
    EnterCemliName = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "ctl00_PlaceHolderMain_nameInput")))
    EnterCemliName.send_keys(CEMLI_NAME)
    CemliCreate = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "ctl00_PlaceHolderMain_createButton")))
    CemliCreate.click()
    if driver.find_elements_by_xpath("//span[contains(@class,'ms-error') and .//text()='The specified name is already in use.']"):
        print("Entered Cemli Page Already Exists \n https://geappliances.sharepoint.com/sites/erpdevops/SitePages/"+CEMLI_NAME+".aspx")
        sys.stdout.flush()
    else:
        EditSourceClick = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.Markup.Html.Menu.Html.EditSource-Large")))
        EditSourceClick.click()
        HTMLCode = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "PropertyEditor")))
        HTMLCode.clear()
        x = '<table width="1313" cellspacing="0" height="954" class="ms-rteTable-0"> <tbody> <tr class="ms-rteTableEvenRow-0"> <td class="ms-rteTableEvenCol-0" style="width: 50%;"> <table width="100%" cellspacing="0" class="ms-rteTable-default"> <tbody> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default" colspan="2" style="width: 50%;"> <strong class="ms-rteFontSize-3">INFORMATION</strong></td> </tr> <tr class="ms-rteTableOddRow-default"> <td valign="top" class="ms-rteTableEvenCol-default" style="width: 13%;"> <strong>JIRA Ticket</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> <br/> </td> </tr> <tr class="ms-rteTableEvenRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Name</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> <br/> </td> </tr> <tr class="ms-rteTableEvenRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Description</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> <br/> </td> </tr> <tr class="ms-rteTableOddRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Source</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default"> Source CI: <br/>Database Table: <br/>Directory: <br/>File: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td valign="top" class="ms-rteTableEvenCol-default"> <strong>Integration</strong><br/></td> <td valign="top" class="ms-rteTableOddCol-default">Integration CI: <br/>Objects: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Target</strong></td> <td class="ms-rteTableOddCol-default">Target CI: <br/>API: <br/>Database Table: <br/>Directory: <br/>File: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Dependencies</strong></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>External</strong></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Contacts</strong></td> <td class="ms-rteTableOddCol-default">Source: <br/>Integration: <br/>Target: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Reprocessing</strong></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Regression Testing</strong><br/></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Support Transition Status</strong><br/></td> <td class="ms-rteTableOddCol-default">MD2060: <br/>Code Walk Through: <br/>CEMLI Page: <br/>Final Sign Off: <br/> </td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>Sample Payload</strong><br/></td> <td class="ms-rteTableOddCol-default">?</td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default"> <strong>non-PROD Details</strong><br/></td> <td class="ms-rteTableOddCol-default">dev: <br/>tst: <br/>qa/stg: <br/></td> </tr> </tbody> </table> </td> <td class="ms-rteTableOddCol-0" style="width: 50%;"> <table width="100%" cellspacing="0" class="ms-rteTable-default"> <tbody> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default" colspan="2" style="width: 50%;"> <strong><span><span><strong class="ms-rteFontSize-3">FAILURE MODES<br/></strong></span></span></strong></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default" style="width: 5%;">#1</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#2</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#3</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#4</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#5</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#6</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#7</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableEvenRow-default"> <td class="ms-rteTableEvenCol-default">#8</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableOddRow-default"> <td class="ms-rteTableEvenCol-default">#9</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> <tr class="ms-rteTableFooterRow-default"> <td class="ms-rteTableFooterEvenCol-default" rowspan="1">#10</td> <td class="ms-rteTableOddCol-default">Object: <br/>CI: <br/>Message: <br/>Action: <br/></td> </tr> </tbody> </table> </td> </tr> </tbody></table><p> <br/></p>'
        HTMLCode.send_keys(x)
        time.sleep(3)
        HTMLCodeSubmit = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "sourcedialog_okbutton")))
        HTMLCodeSubmit.click()
        Checkin = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.EditAndCheckout.Checkout-SelectedItem")))
        Checkin.click()
        CheckinContinue = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "statechangedialog_okbutton")))
        CheckinContinue.click()
        ContentType = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select an option']")))
        ContentType.click()
        ContentTypeSelection = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='CEMLIs']")))
        ContentTypeSelection.click()
        Track = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select options']")))
        Track.click()
        TrackSelection = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='Shared']")))
        TrackSelection.click()
        Module = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select options']")))
        Module.click()
        ModuleSelection = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='Sourcing']")))
        ModuleSelection.click()
        JustClick = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//label[contains(@class,'ms-Label ReactFieldEditor-fieldTitle') and .//text()='Description']")))
        JustClick.click()
        Description = driver.find_elements_by_xpath("//input[@placeholder='Enter value here']")[1].send_keys(CEMLI_NAME)
        time.sleep(2)
        Platform = driver.find_elements_by_xpath("//span[contains(@class,'ms-Dropdown-title ms-Dropdown-titleIsPlaceHolder') and .//text()='Select options']")[1].click()
        time.sleep(2)
        PlatformSelection = driver.find_elements_by_xpath("//span[contains(@class,'ms-Dropdown-optionText dropdownOptionText') and .//text()='SOA']")[0].click()
        time.sleep(3)
        DropdownClose = driver.find_elements_by_xpath("//span[contains(@class,'ms-Dropdown-caretDown') and .//text()='']")[5].click()
        time.sleep(2)
        SaveCEMLI = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-Button-label') and .//text()='Save']")))
        SaveCEMLI.click()
        Edit = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(@class,'ms-promotedActionButton-text') and .//text()='Edit']")))
        Edit.click()
        Checkout = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "Ribbon.EditingTools.CPEditTab.EditAndCheckout.Checkout-SelectedItem")))
        Checkout.click()
        Finalcheckout = WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "statechangedialog_okbutton")))
        Finalcheckout.click()
        print ("CEMLI Sheet has been created succesfully please find the below CEMLI Sheet Link \n ")
        sys.stdout.flush()
        print("https://geappliances.sharepoint.com/sites/erpdevops/SitePages/"+CEMLI_NAME+".aspx")
        sys.stdout.flush()
except NoSuchElementException:
    print ("Error Message :- Element Not Found and Execution got Failed \n Reason :- Element we are trying to access is not found or Slow Network Connection Time Out")
    sys.stdout.flush()
except StaleElementReferenceException:
    print ("Error Message :- Stale Element reference \n Reason :- Stale Element means an old element or no longer available element")
    sys.stdout.flush()
except:
    print ("Error Message :- Some Thing Went Wrong or Flow Terminated abruptly  \n")
    sys.stdout.flush()
finally:
    print ("***************** CEMLI Creation Completed *****************")
    sys.stdout.flush()
		

			
			

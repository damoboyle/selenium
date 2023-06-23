"""
23 December 2022
Damian O'Boyle

This code is designed to automate the daily manual STD/HIV download from SPHL OpenELIS

Prerequisites:
    Access to the Network Drive                 ASAP Request
    Access to OpenELIS (SHPL Query Portal)      ASAP Request

    Python Installed            ITSD Ticket
        Selenium Package          (Download via command prompt "py -m pip install selenium")
        ElementTree Package       (Download via command prompt "py -m pip install elementpath")
        win32com Package          (Download via command prompt "py -m pip install pywin32")
        decouple Package          (Download via command prompt "py -m pip install python-decouple")
        
    Login Credentials   Store your OpenELIS login credentials in a file named ".env" on your personal Drive to maintain security
                        DO NOT 'name' the file or put any text before ".env" in the filename or the system will not be able to read the file
                        The contents of this file should resemble the eg. below (change the X's with your own username/password)
                            USERNAME=XXXXXX
                            PASSWORD=XXXXXXXX
                        Ensure the path to this file matches that on line 61 below (DO NOT CHANGE THE CODE - MOVE YOUR FILE TO MATCH)
    Download Folder     Ensure that chrome is set to download files into the system Download folder (Must match line 67)
                        
Automation Requirements:    Only one preson at a time should have the automation process active on their machine (Damian O'Boyle as of 01/04/2023)
                            The python script file must be stored on your local Drive (Leave a working copy available on the Network Drive)
    Batch File              Create a bacth file (.bat) using notepad and containing the following eg. text 
                              py C:\\User\XXXXXX\Documnets\SPHL_OE2_STD_HIV.py
    Task Scheduler          Search for Task Scheduler and open the program
                            Select Create Basic Task...
                            Complete Name, Description <Next>
                            Select Daily <Next>
                            Set Time (Currently 7:00:00 AM) <Next>
                            Select Start a Program <Next>
                            Insert the path to yout batch (.bat) file <Next>
                            <Finish>
"""
import os
import sys
import time
import shutil
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from decouple import Config, RepositoryEnv
from xml.etree import ElementTree
from win_email import email

#Time and Date
hour = " 00:00"  #(REQUIRED)#Leave single blank space before characters
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
stamp = str(yesterday) + hour + ".." + str(today) + hour

#Login Credentials  
credentials = Config(RepositoryEnv("C:/.env"))  #Used to obfuscate credentials rather than hardcoding
username = credentials.get('USERNAME')
password = credentials.get('PASSWORD')

#File Paths/URLs
folder = "C:/Lab Reports/SPHL HIV_STD/"
download = "C:/Users/" + username + "/Downloads/"
query = "C:/Lab Reports/SPHL HIV_STD/Query/DataView.xml"
new = "HIV STD Results " + str(today).replace('-', '') + ".xlsx"
chromedriver = "C:/chromedriver.exe"
openelis = "https://openelis.____.__.___/openelis/OpenELIS.html"
    
#Email Lists
email_list = "angela.mckee@email.com; damian.oboyle@email.com"
error_list = "angela.mckee@email.com; damian.oboyle@email.com"
#Email Subject
subject = "Daily SPHL STD HIV Results"
noRecord = "No Records in Daily SPHL STD HIV Results"
error = "ERROR in Daily SPHL STD HIV Results"
#Email Body
success = "Please find the most recent STD results from OE2 in the results folder." 

#Start Driver/Enable Actions/Set Wait Time
driver = webdriver.Chrome(service=Service(executable_path=chromedriver))
actions = ActionChains(driver)
wait = WebDriverWait(driver, 30)

def dailyDataPull():
    try:#Log into Query Application
        driver.get(openelis)
        driver.find_element("name", "username").send_keys(username)
        pwd = driver.find_element("name", "j_password")
        pwd.send_keys(password)
        pwd.submit()
        
        #Wait until Webpage/Report button has loaded
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[text()='Report']"))) 
        
    except:#Handles failure to log in
        email(error_list, error, "Failed to log in to OpenELIS")
        exit()
    
    try:
        #Naviagte to Query setup/input
        driver.find_element(By.XPATH, "//div[text()='Report']").click()
        actions.move_to_element(driver.find_element(By.XPATH, "//td[text()='Data Export']")).perform()
        driver.find_element(By.XPATH, "//td[text()='Data View']").click()

        #Wait for Webpage/Upload utility to load
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
        
        try:#Modify Query File (Reset Dates)
            tree = ElementTree.parse(query)
            tree.find('.//_sample.releasedDate').text = stamp
            tree.write(query, encoding="utf-8", xml_declaration=True)
            
        except:#Handles XML query update failure
            email(error_list, error, "Failed to modify Query in 'DataView.xml'")
            exit()
            
        #Upload Query File
        driver.find_element(By.XPATH, "//input[@type='file']").send_keys(query)
        
    except:#Handles failed Query Upload
        email(error_list, error, "Failed to upload Query")
        exit()

    #Set Dates*Does not actually change the dates
    ###NEEDED TO STALL/WAKE THE BROWSER IN ORDER TO EXECUTE QUERY###
    driver.find_element(By.XPATH, "(//input[@class='GDTY2-NBHV'])[7]").send_keys(str(yesterday) + hour)
    driver.find_element(By.XPATH, "(//input[@class='GDTY2-NBHV'])[8]").send_keys(str(today) + hour)
    
    try:#Execute Query
        driver.find_element(By.XPATH, "//div[text()='Execute Query']").click()
        
        #Wait for Webpage/Checkboxes to load
        wait.until(EC.presence_of_element_located((By.XPATH, "(//div[text()='Select All'])")))

        #Select All Fields
        driver.find_element(By.XPATH, "(//div[text()='Select All'])[1]").click()
        driver.find_element(By.XPATH, "(//div[text()='Select All'])[3]").click()
        
    except:#Handles 'no records found' Exception
        if driver.find_element(By.XPATH, "(//div[@class='GDTY2-NBFVB'])[2]").text == "No records found":
            email(error_list, noRecord, "There were no results with a sample_released_date " + str(yesterday))
            exit()
        else:#Handles Error Selecting Variables
            email(error_list, error, "Error  Selecting Variables")
            exit()
    
    #Recursively run the report in order to catch download errors and repeat
    runReport()
        
    #Send Email/Close Browser
    email(email_list, subject, success) #calls seperate script - win_email
    driver.quit()
    
def runReport():
    try:
        driver.find_element(By.XPATH, "//div[text()='Run Report']").click()

        #Wait for file to Generate/Download
        wait.until(EC.text_to_be_present_in_element((By.XPATH, "(//div[@class='GDTY2-NBEXB'])[2]"), "Generated file "))
        
    except:#Handles Failed Download Query Errors
        email(error_list, error, "Query/Download Failed (Internal Error)")
        exit()
            
    try:#Get File name
        filename = driver.find_element(By.XPATH, "(//div[@class='GDTY2-NBEXB'])[2]").text
        filename = "data" + filename.strip('Generated file')        #Strip function removes the word "data" from beginning of filename (adding it back)
    except:#Handles filename error
        email(error_list, error, "Error getting filename")
        exit()
    
    #Loop Block to catch failed downloads
    count = 0
    while count < 5:                                            #Loop sleep/wait
        if os.path.exists(download + filename):
            os.rename(download + filename, download + new)      #Rename Downloaded File
            shutil.move(download + new, folder + new)           #Move File
            count = 5                                           #Loop Break
        else:
            time.sleep(10)                                      #Sleep allows file to fully download
            count += 1
            if count == 5:
                runReport()     #Recursive call to download file again if file was not found after 20 seconds

if __name__ == '__main__':
    dailyDataPull()
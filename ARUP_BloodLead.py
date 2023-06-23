import datetime

#Date Values
today = datetime.datetime.now()
day = today.strftime("%a")
now = today.strftime("%H%M%S")
todate = today.strftime("%m%d%y")

#File Date
if day == "Sun" or day == "Mon":
    exit()
elif day == "Tue":
    date = (today - datetime.timedelta(days=5)).strftime("%m%d%Y")
else:
    date = (today - datetime.timedelta(days=3)).strftime("%m%d%Y")

import os
import time
import pandas   
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from decouple import Config, RepositoryEnv
from win_email import email

#Login Credentials  
credentials = Config(RepositoryEnv("C:/.env"))
username = credentials.get('USERNAME')
password = credentials.get('ET_PASSWORD')

#File Paths/URLs
folder = "C:/HL7 Messaging/ELR_Load_Files/ARUP/"
backup = "C:/HL7 Messaging/ELR_Load_Files_bkp/ARUP/"
download = "C:/Users/" + username + "/Downloads/"
new = "arup" + todate + "_" + now + ".txt"
chromedriver = "S:/Lab Reports/SPHL HIV_STD/Query/chromedriver.exe"
ARUPConnect = "https://connect.aruplab.com/SecureFileTransfer/"

#Email Lists
email_list = "angela.mckee@email.com; damian.oboyle@email.com"
error_list = "angela.mckee@email.com; damian.oboyle@email.com"
#Email Subject
subject = "Daily ARUP Lead Results"
noRecord = "No Records in " + subject
error = "ERROR in " + subject
#Email Body
success = "The daily blood lead results from ARUP Connect have been placed in the ARUP results folder."

#Start Driver/Enable Actions/Set Wait Time
driver = webdriver.Chrome(service=Service(executable_path=chromedriver))
actions = ActionChains(driver)
wait = WebDriverWait(driver, 30)

try:#Log into ARUPConnect
    driver.get(ARUPConnect)

    user = driver.find_element("name", "username")
    user.send_keys(username)
    user.submit()

    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='okta-signin-password']")))

    pwd = driver.find_element("name", "password")
    pwd.send_keys(password)
    pwd.submit()

except:#Handles failure to log in
    email(error_list, error, "Failed to log in to ARUP Connect")
    exit()
    
try:#Download File
    wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(), '" + date + "')]")))

    file = driver.find_element(By.XPATH, "//a[contains(text(), '" + date + "')]")
    file.click()
    file = file.text.split(' ')[0]

    #Wait for File to Download
    while not os.path.exists(download + file):
        time.sleep(1)
        
except:#Handles failure to download file
    email(error_list, error, "Failed to Download File from ARUP Connect")
    exit()
    
try:#Convert File
    head_fix = ('','','Lead Report','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','','')
    data_xls = pandas.read_excel(download + file, 'Sheet1', index_col=0, dtype=str)
    data_xls.to_csv(folder + new, encoding='utf-8', header = head_fix)
 
except:#Handles failure to convert file
    email(error_list, error, "Failed to Convert file from XLS to TXT")
    exit()
    
try:#Move/Backup Files
        os.remove(download + file)
        shutil.copy(folder + new, backup + new)
        
except:#Handles failure to move/backup file
    email(error_list, error, "Failed to backup file")
    exit()
        
email(email_list, subject, success) #calls seperate script - win_email
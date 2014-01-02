from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time
import codecs
from time import strftime
import SendKeys
import psutil
import os
import subprocess
import win32clipboard

## SETTINGS ############

filepath = "****INPUT PATH******"  #double backslashes required
usingFirefox = True
username = "****"
password = "****"

########################

if(usingFirefox):
    browser = webdriver.Firefox()
else:
    browser = webdriver.Ie()
    
browser.get("https://www.iii.co.uk/auth/reg/wealth/?orig_url=/reg/wealth/&lo=1&type=shares_and_funds&orig_server=www.iii.co.uk&orig_method=POST&folder=MyPortfolio&") # Load page

time.sleep(2)
#webdriver.switch_to.window(webdriver.window_handles.last);
time.sleep(2)

time.sleep(0.5)
elem = browser.find_element_by_name("III_username") # Find the query box
elem.clear()
time.sleep(0.2)
elem.send_keys(username)
time.sleep(0.2)
elem = browser.find_element_by_name("III_password")
elem.clear()
time.sleep(0.2)
elem.send_keys(password)
time.sleep(0.2)
elem.send_keys(Keys.ENTER)
time.sleep(10) # Let the page load

# make directory
dirc = filepath+"\\X"+strftime("%Y")+"\\"+strftime("%m")
if not os.path.exists(dirc):
    os.makedirs(dirc)

## save in notepad

SendKeys.SendKeys("^a") 
SendKeys.SendKeys("^c") 

time.sleep(0.1)
SendKeys.SendKeys("%{F4}")

subprocess.Popen(["notepad.exe"])
time.sleep(2)
SendKeys.SendKeys("^v")
time.sleep(2)

SendKeys.SendKeys("^s")
time.sleep(2)

strn = dirc+"\\"+strftime("%y%m%d")+".txt"

SendKeys.SendKeys(strn)
SendKeys.SendKeys("{ENTER}")
time.sleep(1)
SendKeys.SendKeys("%{F4}")

if(not usingFirefox):

    PROCNAME = "IEDriverServer.exe"

    for proc in psutil.process_iter():
        if proc.name == PROCNAME:
            proc.kill()



file = "***** OUTPUT PATH ********"
# instead of open(file,'w') do
f = codecs.open(file, encoding='utf-8',mode='w+')

win32clipboard.OpenClipboard()
c = win32clipboard.GetClipboardData()
f.write(c.decode('cp1252'))

# Date-Created: 7/28/2020
# Original Creator: Franklin
# Last Edited: 7/28/2020
# Current Editors: Franklin

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import smtplib
from pathlib import Path
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import holidays
import time
import os
import sys

"""
Documentation Start 

************************************************************************************************        
************************************************************************************************        
DO NOT upload this file to "The Manual" google doc without removing credentials, please follow the instructions listed in "The Manual"
************************************************************************************************    
************************************************************************************************ 

@general: Every single function and webdriver boilerplate use some kind of reactive programming. In other words,
before executing the majority of code it will wait until something either loads up or is available. This is due to
the non-instantaneous nature of the internet. If new code is implemented that requires data to load it is recommended
you use selenium's WebDriverWait(driver, timeout).until(element) functionality for web related events or set a manual
sleep timer if necessary like the following code block

    timeout = 0
    while not os.path.exists('BL34C-Bills Overdue From Vendor - Excel Format.xlsx') and timeout != 30:
        timeout = timeout + 1
        time.sleep(1)           

@param webdriver: Allows you to use selenium to traverse your browser.

@class generateExceptionReport(): If you want to add exception handling, please add it to this class. It acts just like a switch
statement and can be reused for any portion of the code that requires error handling. It outputs an error log file and a screenshot
of the page in the location the exception was raised. 

@function closeNonSenseEnergySaverDialogBox(): Sometimes energy cap will load up some energy saver pop-up upon signing in and
it does not leave even if you traverse to an entirely different page. Hopefully, they don't introduce another of this kind so
that you don't have to implement another workaround just to close it but it isn't too bad if you do. If the dialog box ever
disappears for good you can remove the function call at the bottom or just leave it. It'll eventually timeout on its own anyway.

@functions tickboxes() && enterfields(): Energy cap will remember he settings you selected for any given report on a user to
user basis. The functions are written so that if someone, new or withstanding, would like to modify the code for their own
energy cap account it can account for freshly made or saved entries. That's why it clears entries and checks if the boxes
are ticked before proceeding.

@function mailit(): HTML is sick. Gmail doesn't natively support tables so the options you're left with is copy and pasting one
from google docs or using html to format the table into the email using pandas dataframes.

@param .format(df.to_html()): .format() takes a curly braced parameter inside of quotes and fills it in. In this case, it takes
the entire html formatting pandas spits out for the datatable and fills it in where {0} is.

@function cleanup(): Always make sure to quit out of the webdriver to save memory and also to just close out the web browser
the idea of this program is to make it run and leave no trace of it ever having been ran. It's simply more convenient to the end
user to not have to close out some dialog box.

@param BooleanHoliday: This is to save coworkers the headache of receiving an email on their break. Boolean Holiday also sounds
like a great band name.

Documentation End
"""



driver = webdriver.Chrome("C:\\Users\\frank\\Downloads\\chromedriver_win32\\chromedriver.exe")
driver.set_page_load_timeout(60)

now = datetime.now()
date = now.strftime("%m-%d-%Y")
booleanHoliday = date in holidays.US()
mailingList = ["human emails here, import from file"]
username = "human email here"
password = "human password here"
datasource = "organization here"

class IncorrectUrlException(Exception):
    pass

class generateExceptionReport(object):
    def generateExceptionReport(self, exceptionName):
        home = str(Path.home())
        print("There has been an error, generating a log...")
        print("Creating an error log directory if it doesn't exist....")
        try:
            os.mkdir(home + "\BillOverDueErrorLogs\\")
        except FileExistsError:
            print("It does exist, moving along...")
        method_name = 'exception_' + exceptionName
        method = getattr(self, method_name, lambda: 'Invalid')
        return method()

    def exception_IncorrectUrlException(self):
        home = str(Path.home())
        os.mkdir(home + "\ErrorLogs\\")
        path = home + "\BillOverDueErrorLogs\IncorrectUrlException" + now.strftime("%m-%d-%Y-%I-%M-%S")
        driver.get_screenshot_as_file(path + ".png")
        errorLog = open(path + ".txt", "w+")
        errorLog.write("It seems we've run into an IncorrectUrlException \n \n")
        errorLog.write("This was the webpage you were sent to: " + driver.current_url + " \n")
        errorLog.write("Make sure your credentials don't need to be reset and that the url hasn't changed.")
        print("Check: " + path + " , there should be a log folder on it")
        time.sleep(3)
        driver.quit()
        sys.exit()


    def exception_TimeoutException(self):
        home = str(Path.home())
        path = home + "\BillOverDueErrorLogs\TimeoutException" + now.strftime("%m-%d-%Y-%I-%M-%S")
        driver.get_screenshot_as_file(path + ".png")
        errorLog = open(path + ".txt", "w+")
        errorLog.write("It seems we've run into a TimeoutException trying to request the bill report url \n \n")
        errorLog.write("This was the webpage you were sent to: " + driver.current_url + " \n")
        errorLog.write("It's possible the url for the bill report is currently somewhere else or the page never loaded.")
        print("Check: " + path + " , there should be a log folder on it")
        time.sleep(3)
        driver.quit()
        sys.exit()


def starter(username, password, datasource, date):
    timeout = 5
    try:
        target = "https://my.energycap.com/app/login?targetHost=app.energycap.com"
        driver.get(target)
        WebDriverWait(driver, timeout).until(EC.url_to_be(target))
        if not driver.current_url == "https://my.energycap.com/app/login?targetHost=app.energycap.com" or driver.current_url == "https://my.energycap.com/":
            print("Something's amiss...")
            raise IncorrectUrlException("You're not in the right place.")
    except IncorrectUrlException:
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("IncorrectUrlException")


    usernameField = driver.find_element_by_id("usernameInput")
    usernameField.clear()
    usernameField.send_keys(username)

    passwordField = driver.find_element_by_id("passwordInput")
    passwordField.clear()
    passwordField.send_keys(password)

    datasourceField = driver.find_element_by_id("datasourceInput")
    datasourceField.clear()
    datasourceField.send_keys(datasource)

    loginButton = driver.find_element_by_id("loginButton")
    loginButton.submit()

    try:
        BillReport = "https://app.energycap.com/app/manage/report/27350"
        driver.get(BillReport)
        element = EC.presence_of_element_located((By.CSS_SELECTOR, "#manageViewTitle"))
        WebDriverWait(driver, timeout).until(element)
    except TimeoutException:
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("TimeoutException")



def closeNonSenseEnergySaverDialogBox():
    timeout = 5
    try:
        element = EC.presence_of_element_located((By.ID, "CloseBtn_icon"))
        WebDriverWait(driver, timeout).until(element)
        if (driver.find_element_by_id("CloseBtn_icon")):
            element = driver.find_element_by_id("CloseBtn_icon")
            element.click()
    except TimeoutException:
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("TimeoutException")


def enterFields():
    timeout = 5
    try:

        element = EC.presence_of_element_located((By.XPATH,
                                                  "/html/body/ec-app/div/ec-manage-report/ec-manage-view/div/div/div/div/div/section[2]/ec-filters/filters/div/div/div/div[2]/form/table[1]/tbody/tr/td[3]/filter-double/input"))
        WebDriverWait(driver, timeout).until(element)
        elementfound = driver.find_element_by_xpath(
            "/html/body/ec-app/div/ec-manage-report/ec-manage-view/div/div/div/div/div/section[2]/ec-filters/filters/div/div/div/div[2]/form/table[1]/tbody/tr/td[3]/filter-double/input")
        elementfound.clear()
        elementfound.send_keys("3")
    except TimeoutException:
        print("HERE 0")
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("TimeoutException")


def tickBoxes():
    print("Checking to see if the 'Vendor Name' box has been ticked...")
    timeout = 10
    try:
        element = EC.presence_of_element_located((By.ID,
                                                  "checkboxVendorName"))
        WebDriverWait(driver, timeout).until(element)
        elementfound = driver.find_element_by_css_selector(
            "#checkboxVendorName > input")
        actualCheckBox = driver.find_element_by_xpath(
            "//*[@id='checkboxVendorName']");
        elementfoundClass = elementfound.get_attribute("class")
        if (elementfoundClass == "custom-control-input ng-pristine ng-untouched ng-valid ng-empty"):
            print("It hasn't been, ticking now...")
            actualCheckBox.click()
        else:
            print("It has been, moving on...")
    except TimeoutException:
        print("If the timeout exception for enterFields went off, the page probably didn't load.")
        print(
            "If it didn't, check tickBoxes. There's most likely a more nuanced problem taking place with the pathing.")

    try:
        element = EC.presence_of_element_located((By.XPATH,
                                                  "//select[@id='operatorVendor Name']/option[4]"))
        WebDriverWait(driver, timeout).until(element)

        driver.find_element_by_xpath("//select[@id='operatorVendor Name']/option[4]").click()
    except TimeoutException:
        print("HERE 1")
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("TimeoutException")

    try:
        element = EC.presence_of_element_located((By.XPATH,
                                                  "/html/body/ec-app/div/ec-manage-report/ec-manage-view/div/div/div/div/div/section[2]/ec-filters/filters/div/div/div/div[2]/form/table[2]/tbody/tr[5]/td[3]/filter-string/input"))
        WebDriverWait(driver, timeout).until(element)

        textFieldFound = driver.find_element_by_xpath(
            "/html/body/ec-app/div/ec-manage-report/ec-manage-view/div/div/div/div/div/section[2]/ec-filters/filters/div/div/div/div[2]/form/table[2]/tbody/tr[5]/td[3]/filter-string/input")
        textFieldFound.clear()
        textFieldFound.send_keys("Direct Energy")
    except TimeoutException:
        print("HERE 2")
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("TimeoutException")


def download():
    timeout = 5
    try:
        print("Attempting to click on drop down...")
        element = EC.presence_of_element_located((By.XPATH,
                                                  "//*[@id='downloadDropdown_dropdown_toggle_button']"))
        WebDriverWait(driver, timeout).until(element)
        elementfound = driver.find_element_by_xpath(
            "//*[@id='downloadDropdown_dropdown_toggle_button']")
        elementfound.click()
    except TimeoutException:
        print("failed to click on button dropdown, throwing exception...")
        print("Timed out waiting for page to load")
    try:
        print("Attempting to click on drop down item...")
        element = EC.presence_of_element_located((By.XPATH,
                                                  "//*[@id='downloadReportItem_Excel data only_label']"))
        WebDriverWait(driver, timeout).until(element)
        downloadfound = driver.find_element_by_xpath(
            "//*[@id='downloadReportItem_Excel data only_label']")
        downloadfound.click()
    except TimeoutException:
        print("failed to click on button dropdown item, throwing exception...")
        errorHandle = generateExceptionReport()
        errorHandle.generateExceptionReport("TimeoutException")


def mailIt(recieverName):
    print("Writing up emails...")
    timeout = 0
    while not os.path.exists(
            r'C:\Users\frank\Downloads\BL34C-Bills Overdue From Vendor - Excel Format.xlsx') and timeout != 30:
        timeout = timeout + 1
        time.sleep(1)

    now = datetime.now()
    date = now.strftime("%m-%d-%Y")
    df = pd.read_excel(r'C:\Users\frank\Downloads\BL34C-Bills Overdue From Vendor - Excel Format.xlsx', skiprows=5,
                       skipfooter=3)
    fromaddr = "billsoverduetemple@gmail.com"
    toaddr = recieverName

    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr

    msg['Subject'] = "Bills Due Report " + date

    body = """\
<html>
  <head></head>
  <h3>Attached is the excel file from Energy Cap detailing the bills overdue and below is a preview of bills that are due. <br> <br> <br> <br> </h3>
  <h4>If there are any inconsistencies with the information in the table, please notify the editor.</h4>
  <body>
    {0}
  </body>
</html>
""".format(df.to_html())

    msg.attach(MIMEText(body, 'html'))

    filename = 'BL34C-Bills Overdue From Vendor - Excel Format.xlsx'
    attachment = open(r'C:\Users\frank\Downloads\BL34C-Bills Overdue From Vendor - Excel Format.xlsx', "rb")

    p = MIMEBase('application', 'octet-stream')

    p.set_payload((attachment).read())

    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(p)

    s = smtplib.SMTP('smtp.gmail.com', 587)

    s.starttls()

    s.login(fromaddr, "{Password Here}")

    text = msg.as_string()
    s.sendmail(fromaddr, toaddr, text)
    s.quit


def cleanup():
    os.remove(r"C:\Users\frank\Downloads\BL34C-Bills Overdue From Vendor - Excel Format.xlsx")
    driver.quit()
    print("Daily task successfully completed!")




# Artificial New Years below for testing
# forcedHoliday = "01-01-2021" in holidays.US();

if not booleanHoliday:
    print("Today's date is: ")
    print(date)
    print("Beginning bill report script...")
    starter(username, password, datasource, date)
    closeNonSenseEnergySaverDialogBox()
    enterFields()
    tickBoxes()
    download()
    for staffMember in mailingList:
        mailIt(staffMember)
    cleanup()
else:
    print("Today's date is: ")
    print(date)
    print("Hey! It's a holiday! Leave them alone!")



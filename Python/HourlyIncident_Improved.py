import time
import pyautogui as gui
from time import sleep as sl
import getpass, os, datetime
from datetime import date, timedelta
from selenium import webdriver 
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, ElementNotVisibleException

import win32com.client, getpass
from PIL import ImageGrab
import shutil, schedule
import getpass


win32c = win32com.client.constants
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6)
messages=inbox.Items


path = r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Input'

class HourlyIncident:
    def __init__(self,log):
        self.log = log

    def clearfiles(self,log):
        try:
            paths = [r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Input', r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Input\Processed', r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Output']
            for pth in paths:

                files = [fil for fil in os.listdir(pth) if os.path.isfile(os.path.join(pth, fil))]
                for file in files:
                    try:
                        os.unlink(os.path.join(pth, file))
                        print(f"{file} has been deleted")
                        sl(1.5)
                    except Exception as e:
                        print(str(e))
                        # raise e
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def getEmail(self,subject,log):
        # path = 'D:\\2020\\Python\\9Mobile\\SKYAvailability\\Input\\'
        try:
            print('Checking inbox', end = ' ')
            sl(2)
            checker = 0
            for dot in range(len('...')):
                print('...'[dot], sep = ' ', end = ''); sl(1)
            print('')
            for message in messages:
                if subject in message.Subject and message.SenderEmailAddress == 'nigeriarnocautin@ms.huawei.com':
                    print(f'Mail found! Sender is {message.SenderEmailAddress}')
                    attachments = message.Attachments
                    attachment = attachments.Item(1)
                    attachment.SaveAsFile(os.path.join(path,f"{str(attachment)}"))
                    print(f"Attachment {attachment} Downloaded Successfully")
                    sl(2)
                    try:
                        message.Delete()
                        print('Mail Deleted')
                    except:
                        pass
                    checker += 1
                    sl(1)
            if checker > 0:
                return True
            else:
                return False
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def generateImageFile(self,log):
        try:
            createdImgs = []
            processedPath = r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Input\Processed'
            imgPath = r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Output'
            print('Got here02')
            excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
            #excel = win32com.client.DispatchEx("Excel.Application")
            for file in os.listdir(processedPath):
                if 'OBW Network Status Report ' in file and '_-_' in file and not '~$' in file:
                    lastLineNum = int(file[file.index('_-_') + 3:file.index('.xlsx')])
                    wb = excel.Workbooks.Open(os.path.join(processedPath,file))
                    ws = wb.Worksheets['Action in Progress']
                    ws.Range(ws.Cells(1,1),ws.Cells(lastLineNum,9)).CopyPicture(Format= win32c.xlBitmap)
                    img = ImageGrab.grabclipboard()
                    imgFile = os.path.join(os.path.join(imgPath,f'{file[:file.index("_-_")]}.jpg'))
                    createdImgs.append(f'{file[:file.index("_-_")]}.jpg')
                    img.save(imgFile)
                    wb.Close(True)
            return createdImgs
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def launchDriver(self,log):
        try:
            #If chromedriver is not added to path, replace below path with the absolute path to chromedriver in your computer 

            # driver = webdriver.Chrome('D:\\Reporting Python\\chromedriver.exe')
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument('--ignore-certificate-errors')
            driver = webdriver.Chrome(options=chrome_options)

            driver.get("https://web.whatsapp.com/") 
            wait = WebDriverWait(driver, 600) 
            return [driver,wait]
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def whatsappBot(self,driver, wait, fileNames,log):
        try:
            file_day = datetime.date.today() - datetime.timedelta(0)
            dd = file_day.strftime('%I %p : %b-%d-%Y')
            # Contact Name
            target = '"Testing Bot"'

            # Caption_text = f'BW Orange Hourly Incident Report for {dd}'

            print(f'Searching for {target} group', end = '')
            for dot in range(3):
                print('...'[dot],sep = ' ', end = ''); sl(1)

            x_arg = '//span[contains(@title,' + target + ')]'
            group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg))) 
            group_title.click() 
            print("Whatsapp Group Found")
            sl(2)
            whatsapp_input = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
            input_box = wait.until(EC.presence_of_element_located(( 
                By.XPATH, whatsapp_input))) 
            print("Preparing to send Message")

            attach_xpath = '//*[@id="main"]/header/div[3]/div/div[2]/div'
            send_file_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span[2]/div/div/span'

            send_caption_text = '/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div[1]/span/div/div[2]/div/div[3]/div[1]/div[2]'
            
            attachment_type = 'img'

            # TODO - ElementNotVisibleException - this shouldn't happen but when would it

            # local variables for x_path elements on browser
            attach_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/div/span'
            send_file_xpath = '/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div'

            if attachment_type == "img":
                attach_type_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[1]/button/span'
            elif attachment_type == "cam":
                attach_type_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[2]/button/span'
            elif attachment_type == "doc":
                attach_type_xpath = '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[1]/div[2]/div/span/div/div/ul/li[3]/button/span'
            
            for fileName in fileNames:
                
                try:
                    Caption_text = f"OBW Hourly Incident Report {fileName[fileName.index('('):fileName.index('.')]}"
                    # open attach menu
                    attach_btn = driver.find_element_by_xpath(attach_xpath)
                    attach_btn.click()
                    print("Preparing to Attach file")
                    # Find attach file btn and send screenshot path to input
                    sl(1)
                    attach_img_btn = driver.find_element_by_xpath(attach_type_xpath)
                    attach_img_btn.click()
                    sl(5)

                    gui.typewrite(f"D:\\2020\\Excel\\BWOrange\\OBWHourlyIncidentReport\\Output\\{fileName}")
                    gui.press('enter')
                    sl(5)
                    # attachment
                    send_caption = driver.find_element_by_xpath(send_caption_text)
                    send_caption.send_keys(Caption_text)

                    # sl(5)
                    # Pic Caption
                    send_btn = driver.find_element_by_xpath(send_file_xpath)
                    send_btn.click()
                    # send_btn.send_keys(Keys.ENTER)
                    # gui.press('enter')

                    print("Message Sent Successfully")
                    self.remove_pic(fileName,log)
                    sl(3)
                except (NoSuchElementException, ElementNotVisibleException) as e:
                    print(str(e))
                    # send_message(driver,(str(e)))
                    # send_message(driver,"Bot failed to retrieve report content, trying again...")
                    raise e
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def HourlyIncidentReport(self,log):
        try:
            macroDr = "D:\\2020\\Excel\\BWOrange\\OBWHourlyIncidentReport\\HourlyIncident.xlsm"
            print('Got here01')
            # xl = win32com.client.DispatchEx("Excel.Application")
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Application.Visible = True
            print("Loading - Hourly Incident Macro ...")
            wb = xl.Workbooks.Open(os.path.abspath(macroDr))
            # xl.Workbooks.Open(os.path.abspath(macroDr))
            print("Hourly Incident Macro ---  Running")
            xl.Application.Run("'HourlyIncident.xlsm'!HourlyIncident")
            # xl.Application.Quit()
            wb.Close(True)
            print("Completed SUCCESSFULLY!!!")
            sl(5)
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def send_message(self,driver,msg,log):
        # whatsapp_msg = driver.find_element_by_class_name('_2S1VP')
        error_msg =driver.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]")
        sl(2)
        error_msg.send_keys(msg)
        error_msg.send_keys(Keys.ENTER)

    def remove_pic(self,fileName,log):
        try:
            path_f = rf"D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Output\{fileName}"
            os.remove(path_f)
            print("File Deleted Successfully, Waiting for Next Report")
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def afterRun(self,driver,wait,log):
        try:
            self.clearfiles(log)
            if self.getEmail('OBW Network Status Report ',log):
                self.HourlyIncidentReport(log)
                imgFiles =  self.generateImageFile(log)
                condition = True if len(imgFiles) > 0 else False
                if condition:
                    # driver = launchDriver()
                    self.whatsappBot(driver, wait, imgFiles,log)
            sl(10)
            self.afterRun(driver,wait,log)
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

    def firstRun(self,log):
        try:
            self.clearfiles(log)
            if self.getEmail('OBW Network Status Report ',log):
                self.HourlyIncidentReport(log)
                imgFiles =  self.generateImageFile(log)
                condition = True if len(imgFiles) > 0 else False
                if condition:
                    launchD = self.launchDriver(log)
                    driver,wait = launchD[0], launchD[1]
                    self.whatsappBot(driver, wait, imgFiles,log)
                    sl(10)
                    self.afterRun(driver,wait,log)
            else:
                self.firstRun(log)
        except Exception as e:
            print(str(e))
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            raise e

def createDirectories():
        basePath = r'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport'
        folders = [r'Input\Processed','Download','System','Output','Log']
        try:
            for folder in folders:
                try:
                    os.makedirs(os.path.join(basePath,folder))
                    print(f'Path created successfully! {os.path.join(basePath,folder)}')
                    sl(1)
                except:
                    pass
        except Exception as e:
            print(str(e))
            raise e

def main(log):
    HourlyIncident.firstRun(log)

if __name__ == '__main__':
    createDirectories()
    with open(rf'D:\2020\Excel\BWOrange\OBWHourlyIncidentReport\Log\{"HourlyIncident_log.txt"}', "a") as log:
        try:
            startTime = time.time()
            HourlyIncident = HourlyIncident(log)
            main(log)
            endTime = time.time()
        except Exception as e:
            endTime = time.time()
            print(f'Error occurred: {str(e)}')
            sl(3)
            print(f"The bot ran for {int((endTime - startTime)/3600)} hrs : {int((endTime - startTime)/60)} mins : {round(endTime - startTime)%60} secs.")
            log.write(str(datetime.datetime.today())[:18] + '\n')
            log.write(str(e) + '\n')
            log.write(str(datetime.datetime.today())[:18] + '\n')
            # pass

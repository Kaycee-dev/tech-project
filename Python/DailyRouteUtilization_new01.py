import pyautogui as gui
import pyautogui
import numpy as np
import cv2
import datetime
from datetime import date, timedelta
import time as sl
import re, glob,socket
import os, os.path
from os import system
import shutil
import win32com.client
from win32com.client import Dispatch

import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait

CoreATCADir = "E:\\os_CORE\\client\\startup_all_global.bat" 
CoreNFVDir = "E:\\U2000_NFV\\client\\startup_all_global.bat"

location = "E:\\2019\\Download\\"
destination = "E:\\2019\\Airtel\\DailyRouteUtilization\\Input\\"

templatesATCA = ["All Routes template 24hrs", "Mobile Office BSC_RNC 24hrs"]
alttemplatesATCA = ["AllRoutes24", "MobileOff24"]
templatesNFV = ["All Route Templates 24hrs NFV", "Mobile Office BSC_RNC 24hrs NFV"]
alttemplatesNFV = ["AllRoute24NFV", "MobOffBSC24NFV"]

# templatesATCAu2020 = ["All Routes template 24hrs","All Routes template", "CCR All routes", "Mobile Office BSC_RNC 24hrs", "Mobile Office BSC_RNC"]
# alttemplatesATCAu2020 = ["AllRoutes24","All Routes template", "CCR All routes", "MobileOff24", "MobileOffice"]
templatesNFVu2020 = ["All Route Templates 24hrs NFV","All Route Templates NFV", "CCR All routes NFV", "Mobile Office BSC_RNC 24hrs NFV", "Mobile Office BSC_RNC NFV"]
alttemplatesNFVu2020 = ["AllRoute24NFV","AllRoutesNFV", "CCR All routes NFV", "MobOffBSC24NFV", "MobOffRNCNFV"]

today = datetime.date.today()
pdate = date.today() - timedelta(1)

def imagesearch(image, precision=0.8):
    im = gui.screenshot()
    #im.save('testarea.png') usefull for debugging purposes, this will save the captured region as "testarea.png"
    img_rgb = np.array(im)
    img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
    template = cv2.imread(image, 0)
    template.shape[::-1]

    res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val < precision:
        return [-1,-1]
    return max_loc

def imagesearch_loop(image, timesample, precision=0.8):
    pos = imagesearch(image, precision)
    while pos[0] == -1:
        print(image+" not found, waiting")
        sl.sleep(timesample)
        pos = imagesearch(image, precision)
    return pos


def deletefiles():
    paths = [location, destination]
    # paths = [dwnattachmnt]
    for pth in paths:

        files = [fil for fil in os.listdir(pth) if os.path.isfile(os.path.join(pth, fil))]
        for file in files:
            try:
                os.unlink(os.path.join(pth, file))
                print("{} has been deleted".format(file))
                sl.sleep(1)
            except Exception as e:
                print(str(e))

def MoveFile(location, destination):
    sl.sleep(3)
    shutil.move(location, destination)

def renamefile(name,location,destination):
    # path = "C:\\Users\\uwx815001\\Desktop\\9TT Capture Rate\\Input File\\Downloads\\"
    file = os.listdir(location)
    ext = 1
    for efile in file:

        if os.path.isfile(os.path.join(location, efile)) and ".tar.gz" not in efile:
            os.rename(os.path.join(location, efile), os.path.join(location, f'{name}.xlsx'))
            print(f"{name} file renamed successfully")
            sl.sleep(5)
            # DailyTT.MoveFile(os.path.join(path, name + '.xlsx'),destination)
            if not os.path.exists(os.path.join(destination, f'{name}.xlsx')):
                MoveFile(os.path.join(location, f'{name}.xlsx'),destination)
                sl.sleep(5)
                print(f"{name} file moved successfully")
            else:

                os.rename(os.path.join(location,  f'{name}.xlsx'), os.path.join(location, f'{name}{ext}.xlsx'))
                while os.path.exists(os.path.join(destination, f'{name}{ext}.xlsx')):
                    ext+=1
                    os.rename(os.path.join(location, f'{name}{ext-1}.xlsx'), os.path.join(location, f'{name}{ext}.xlsx'))
                    print(f"{name}{ext} file renamed successfully")
                    
                MoveFile(os.path.join(location, f'{name}{ext}.xlsx'),destination)
                print(f"{name}{ext} file moved successfully")
                sl.sleep(5)
        elif os.path.isfile(os.path.join(location, efile)) and ".tar.gz" in efile:
            os.rename(os.path.join(location, efile), os.path.join(location, f'{name}.tar.gz'))
            print(f"{name} file renamed successfully")
            sl.sleep(10)
            ExtractZip(name)
            if not os.path.exists(os.path.join(destination, f'{name}.CSV')):
                MoveFile(os.path.join(location, f'{name}.CSV'),destination)
                sl.sleep(10)
                print(f"{name} file moved successfully")
            else:

                os.rename(os.path.join(location,  f'{name}.CSV'), os.path.join(location, f'{name}{ext}.CSV'))
                while os.path.exists(os.path.join(destination, f'{name}{ext}.CSV')):
                    ext+=1
                    os.rename(os.path.join(location, f'{name}{ext-1}.CSV'), os.path.join(location, f'{name}{ext}.CSV'))
                    print(f"{name}{ext} file renamed successfully")
                    sl.sleep(10)
                    
                MoveFile(os.path.join(location, f'{name}{ext}.CSV'),destination)
                print(f"{name}{ext} file moved successfully")
                sl.sleep(10)

def pull_Data():
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\saveRecord.png',0.5)
    gui.doubleClick(pos[0]+10,pos[1])
    sl.sleep(5)
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\HomeImg.png',0.5)
    gui.doubleClick(pos[0]+10,pos[1])
    sl.sleep(4)
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\ReportFile.png',0.5)
    gui.doubleClick(pos[0]+10,pos[1])
    sl.sleep(4)
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\saveFile.png',0.5)
    gui.click(pos[0]+10,pos[1])
    sl.sleep(5)
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\saveOK.png',0.5)
    gui.click(pos[0]+10,pos[1])
    print("Data saved successfully")
    sl.sleep(2)

def Exitapp():  
    sl.sleep(5)  
    gui.hotkey('alt','F4')
    sl.sleep(5)
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\Exitapp.png',0.5)
    gui.click(pos[0]+10,pos[1])
    print("Exiting U2000")
    sl.sleep(5)

def CoreLogin(CoreName,CoreDir):
    
    with open('E:\\2019\\Airtel\\DailyRouteUtilization\\System\\details.txt','r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    curdir = os.getcwd()
    if CoreDir == CoreATCADir:
        os.chdir("E:\\os_CORE\\client\\")
    elif CoreDir == CoreNFVDir:
        os.chdir("E:\\U2000_NFV\\client\\")
    system(f'cmd /c {CoreDir}')
    os.chdir(curdir)
    print(f'waiting for Core{CoreName}.....')
    sl.sleep(6)
    gui.typewrite(username)
    sl.sleep(2)
    gui.press('tab')
    sl.sleep(2)
    gui.typewrite(password)
    sl.sleep(5)
    pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\loginU2000.png',0.5)
    gui.click(pos[0]+10,pos[1])
    print("Login successfully")
    sl.sleep(30)
    if CoreDir == CoreNFVDir:
        gui.hotkey('esc')
        sl.sleep(5)
    else:
        for num in range(counter):
            pos = imagesearch_loop('E:\\2019\\Airtel\\ReportPull\\saveOK.png',0.5)
            gui.click(pos[0]+10,pos[1])
            sl.sleep(2)
    try:
        gui.hotkey('alt','p')
        gui.hotkey('q')
    except:
        try:
            gui.hotkey('alt','p')
            gui.hotkey('q')
        except:
            pass
    sl.sleep(10)


def PullData(TempName,whch):
    if whch == 1:
        pos = imagesearch_loop(f'E:\\2019\\Airtel\\ReportPull\\searchTemplate.png',0.5)
        gui.click(pos[0]+10,pos[1])
        sl.sleep(2)
    else:
        pos = imagesearch_loop(f'E:\\2019\\Airtel\\ReportPull\\NextTemp.png',0.5)
        gui.click(pos[0]+10,pos[1])
        sl.sleep(2)
        gui.hotkey('ctrl', 'a')
        sl.sleep(2)
    
    gui.typewrite(str(TempName))
    gui.press('enter')
    sl.sleep(10)
    pos = imagesearch_loop(f'E:\\2019\\Airtel\\ReportPull\\{TempName}.png',0.5)
    gui.doubleClick(pos[0]+10,pos[1])
    sl.sleep(35)
    pull_Data()
    

def u2020():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    prefs = {'download.default_directory': destination}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    # driver.quit()
    
    with open('E:\\2019\\Airtel\\DailyRouteUtilization\\System\\details.txt','r') as credentials:
        details = credentials.readlines()
        username,password,counter = details[0].strip(),details[1].strip(),int(details[2].strip())

    driver.get("https://10.227.118.133:31943")
    sl.sleep(5)
    driver.refresh()
    sl.sleep(5)
    driver.find_element_by_id("username").send_keys(username)
    sl.sleep(2)
    driver.find_element_by_id("value").send_keys(password)
    sl.sleep(2)
    driver.find_element_by_id("submitDataverify").click()
    sl.sleep(5)

    def performanceQueryOpen():
        pos = imagesearch_loop('E:\\2020\\Excel\\Airtel\\ReportPull\\QueryResult.png',0.5)
        gui.click(pos[0]+10,pos[1])
        sl.sleep(3)
        

    def performanceQueryClose():
        js = 'document.getElementsByClassName("cancel-img")[0].click();'
        driver.execute_script(js)
        sl.sleep(3)
    # driver.find_element_by_xpath("//*[@id="home_menu_list"]/section/div/div/a/div/div[2]").click()
    # sl.sleep(3)
        
    for num in range(len(templatesNFVu2020)):
        if (num == 0 or num%6 == 0):
            performanceQueryOpen()
            PullData(alttemplatesNFVu2020[num],1)
            renamefile(templatesNFVu2020[num],location,destination)
            if num == len(templatesNFVu2020) - 1:
                performanceQueryClose()
        elif num != len(templatesNFVu2020) - 1 and num%5 != 0:
            PullData(alttemplatesNFVu2020[num],0)
            renamefile(templatesNFVu2020[num],location,destination)
        elif num == len(templatesNFVu2020) - 1 or num%5 == 0:
            PullData(alttemplatesNFVu2020[num],0)
            renamefile(templatesNFVu2020[num],location,destination)
            performanceQueryClose()
    driver.quit()


def DailyRouteUtilizationReport():
    macroDr = "E:\\2019\\Airtel\\DailyRouteUtilization\\DailyRouteUtilization.xlsm"
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Application.Visible = True
    print("Loading - Daily Route Utilization ...")
    wb = xl.Workbooks.Open(os.path.abspath(macroDr))
    print("Daily Route Utilization ---  Running")
    xl.Application.Run("'DailyRouteUtilization.xlsm'!DailyAInterface")
    xl.Application.Quit()
    print("Completed SUCCESSFULLY!!!")
    sl.sleep(5)


def main():
    deletefiles()
    for num in range(len(templatesATCA)):
        if (num == 0 or num%6 == 0):
            CoreLogin("CoreATCA",CoreATCADir)
            PullData(alttemplatesATCA[num],1)
            renamefile(templatesATCA[num],location,destination)
            if num == len(templatesATCA) - 1:
                Exitapp()
        elif num != len(templatesATCA) - 1 and num%5 != 0:
            PullData(alttemplatesATCA[num],0)
            renamefile(templatesATCA[num],location,destination)
        elif num == len(templatesATCA) - 1 or num%5 == 0:
            PullData(alttemplatesATCA[num],0)
            renamefile(templatesATCA[num],location,destination)
            Exitapp()

    # for num in range(len(templatesNFV)):
    #     if (num == 0 or num%6 == 0):
    #         CoreLogin("CoreNFV",CoreNFVDir)
    #         PullData(alttemplatesNFV[num],1)
    #         renamefile(templatesNFV[num],location,destination)
    #         if num == len(templatesNFV) - 1:
    #             Exitapp()
    #     elif num != len(templatesNFV) - 1 and num%5 != 0:
    #         PullData(alttemplatesNFV[num],0)
    #         renamefile(templatesNFV[num],location,destination)
    #     elif num == len(templatesNFV) - 1 or num%5 == 0:
    #         PullData(alttemplatesNFV[num],0)
    #         renamefile(templatesNFV[num],location,destination)
    #         Exitapp()
    u2020()
    DailyRouteUtilizationReport()

if __name__ == "__main__":
    strtime = sl.time()
    main()
    endtime = sl.time()
    print(f"The Operation took {round(endtime - strtime)//60} mins : {round(endtime - strtime)%60} secs to complete.")
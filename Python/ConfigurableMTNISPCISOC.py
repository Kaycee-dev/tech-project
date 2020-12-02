import win32com.client
import getpass
import os
import time as sl
import datetime
from datetime import date, timedelta, datetime
import pandas as pd
import time as sl
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import zipfile
import shutil
import schedule

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

global dashboard

subject = "Modify staff details: NG MTN ISPC/ISOC Appraisal Dashboard"
emailList = ['joy.chinenye.ozoudeh@huawei.com','uba.kelechi.michael@huawei.com', 'wangmu.jerome@huawei.com','akinde.adeniyi.femi1@huawei.com']

changeInitiators = {
    'joy.chinenye.ozoudeh@huawei.com':'Joy Chinenye Ozoudeh j84154681',
    'uba.kelechi.michael@huawei.com':'Uba Kelechi Michael WX815001',
    'akinde.adeniyi.femi1@huawei.com':'Akinde Adeniyi Femi a00494151'
}

class UpdateUser:

    def __init__(self, checkRun, curdir,dependentFileDirectory,backupFileDirectory):
        self.checkRun = checkRun
        self.curdir = curdir
        self.dependentFileDirectory = dependentFileDirectory
        self.backupFileDirectory = backupFileDirectory

    def MoveFile(self, location, destination):
        sl.sleep(3)
        shutil.move(location, destination)

    def renamefile(self, name, location, destination):
        file = os.listdir(location)
        extn = 1
        for efile in file:
            if os.path.isfile(os.path.join(location, efile)) and efile == name:
                if not os.path.exists(os.path.join(destination, name)):
                    self.MoveFile(os.path.join(location, name),destination)
                    sl.sleep(5)
                    print(f"{name} file moved successfully")
                else:

                    os.rename(os.path.join(location, name), os.path.join(location, f'Backup0{extn}_{name}'))
                    print(f"Backup0{extn}_{name} file renamed successfully")
                    while os.path.exists(os.path.join(destination, f'Backup0{extn}_{name}')):
                        extn+=1
                        os.rename(os.path.join(location, f'Backup0{extn-1}_{name}'), os.path.join(location, f'Backup0{extn}_{name}'))
                        print(f"Backup0{extn}_{name} file renamed successfully")
                        
                    self.MoveFile(os.path.join(location, f'Backup0{extn}_{name}'),destination)
                    print(f"Backup0{extn}_{name} file moved successfully")
                    sl.sleep(5)

    def startRun(self):
        
        while self.checkRun:
            print("Checking for email...")
            for message in messages:
                # sl.sleep(1)
                if subject in message.Subject and message.SenderEmailAddress in emailList:
                    global messageSender
                    messageSender = message.SenderEmailAddress
                    print('sender is ' + messageSender)
                    sl.sleep(3)
                    body_content = message.body
                    attachments = message.Attachments
                    attachment = attachments.Item(1)
                    if str(attachment) == 'User Information.xlsx':
                        dashboard = 'ISPC'
                    elif str(attachment) == 'MTN  FO_new sheet.xlsx':
                        dashboard = 'ISOC'
                    print(f'dashboard is {dashboard}')
                    # os.system(f"pushd {dependentFileDirectory}")
                    os.chdir(dependentFileDirectory)
                    print(f'Working directory changed from {curdir} to {os.getcwd()}')
                    self.renamefile(str(attachment),dependentFileDirectory,backupFileDirectory)
                    attachment.SaveAsFile(os.path.join(dependentFileDirectory,str(attachment)))
                    os.chdir(curdir)
                    with open("D:\\2020\\Email\\Log\\ISOC_ISPC_log.txt", "a") as log:
                        log.write(str(datetime.today())[:18] + '\n')
                        log.write(f'Change made by {message.SenderEmailAddress} \t {message.Subject}' + '\n')
                        log.write(str(datetime.today())[:18] + '\n')
                    # os.chdir("ISPC_ISOC_DashboardBackup")
                    print(f"{attachment} file updated successfully!")
                    message.Delete()
                    print(f'Mail Deleted!')
                    sl.sleep(10)
                    self.checkRun = False
            sl.sleep(60)
        self.UpdateUserInfoMailer(dashboard)

    def UpdateUserInfoMailer(self,dashboard):
        print("Sending Email in progress....")
        sl.sleep(2)
        exportTime = str(datetime.now())[:19] 
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = 'joy.chinenye.ozoudeh@huawei.com; akinde.adeniyi.femi1@huawei.com ;uba.kelechi.michael@huawei.com'
        mail.Subject = f'USER DETAILS CHANGE NOTIFICATION FOR {dashboard} DASHBOARD'
        # mail.Body = ''
        mail.Body = f'Dear All, \n\n A change was just done on {dashboard} dashboard user details by {changeInitiators[messageSender]} on {exportTime}'
        # attachment  = "D:\\2019\\NineMobile\\AboveOneHour\\Output\\AboveOneHour.xlsx" 
        mail.Send()
        print("Update user email sent successfully!!!")
        self.checkRun = True
        self.startRun()

def main():
    try:
        UpdateUser.startRun()
        # UpdateUserInfoMailer()
        
    except Exception as excp:
        os.chdir(curdir)
        with open("D:\\2020\\Email\\Log\\ISOC_ISPC_log.txt", "a") as log:
            log.write(str(datetime.today())[:18] + '\n')
            log.write(str(excp) + '\n')
            log.write(str(datetime.today())[:18] + '\n')
        pass

if __name__ == "__main__":
    strtime = sl.time()
    curdir = os.getcwd()
    dependentFileDirectory = '\\\\172.16.151.44\\Users\\Administrator\\Desktop\\dependent\\'
    backupFileDirectory = 'ISPC_ISOC_DashboardBackup'
    global checkRun
    checkRun = True

    UpdateUser = UpdateUser(checkRun,curdir,dependentFileDirectory,backupFileDirectory)
    main()
    endtime = sl.time()
    print(f"The Operation ran for {(endtime - strtime)//86400} days : {(endtime - strtime)//3600} hours : {(endtime - strtime)//60} mins: {round(endtime - strtime)%60} secs.")      
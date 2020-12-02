import win32com.client
import os
import time as sl
import datetime
# import cx_Oracle
import pandas as pd
import time as sl
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from zipfile import ZipFile 
import shutil, schedule
# import getpass

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox=outlook.GetDefaultFolder(6)
messages=inbox.Items

paths = ['D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\','D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\excel\\','D:\\2020\\Python\\9Mobile\\SKYAvailability\\Input\\','D:\\2020\\Python\\9Mobile\\SKYAvailability\\Output\\']

def deleteFiles(paths):
    try:
        for path in paths:
            for file in os.listdir(path):
                try:
                    os.unlink(os.path.join(path,file))
                    print(f'{file} file deleted...')
                except Exception as excp:
                    print(f'An error occured: {str(excp)}')
    except Exception as e:
        print(f'An error occured at delete phase: {str(e)}')
        raise Exception(e)

def runAll(subject):
    try:
        path = 'D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\'
        print('checking inbox.......')
        global newName
        newName = ""
        for message in messages:
            # print(message.Subject)
            if subject in message.Subject and message.SenderEmailAddress == 'EMTSOSS@9mobile.com.ng':
                #EMTSOSS@9mobile.com.ng
                print(f'sender is {message.SenderEmailAddress}')
                # body_content = message.body
                attachments = message.Attachments
                attachment = attachments.Item(1)
                if not os.path.exists(os.path.join(path,f"{str(attachment)}")):
                    attachment.SaveAsFile(os.path.join(path,f"{str(attachment)}"))
                else:
                    ext = 1
                    newName = f'{str(attachment)}{ext}'
                    while os.path.exists(os.path.join(path,newName)):
                        ext += 1
                        newName = f'{str(attachment)}{ext}'
                    attachment.SaveAsFile(os.path.join(path,newName))
                    
                print(f'{newName if newName != "" else str(attachment)} attachment Downloaded Successfully')
                message.Delete()
                print('Mail Deleted')
                sl.sleep(2)

                # generate new data 
        for zipFolder in os.listdir(path):
            if 'Regional Availability for GIS' in zipFolder:
                # 
                file_name = (path+zipFolder)

                # opening the zip folder in READ mode 
                with ZipFile(file_name, 'r') as zip: 
                    # printing all the contents of the zip file 
                    zip.printdir() 
                
                    # extracting all the files 
                    print('Extracting all the files now...') 
                    zip.extractall(path)
                    print('Done!')    
                

                # upload to OWS

        try:
            genrateFileandUpload()
            # upload('https://103n-saapp.teleows.com/servicecreator/spl/base_dashboard/bdb_hourly_availability_import_page.spl?','HourlyAvailability.xlsx')
            print('old data updated')
            return True 
        except:
                    pass
    except Exception as e:
        print(f'An error occured at runAll phase: {str(e)}')
        raise Exception(e)

def genrateFileandUpload():

    for file in os.listdir('D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\excel'):
        if 'Reg Availa for GIS' in file:

            try:
                df  = pd.read_excel(f"D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\excel\\{file}", "2G CA",skiprows = 0)
                df = df.set_index("Time")
                # cl = df.columns.tolist()

                df2  = pd.read_excel(f"D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\excel\\{file}", "3G CA",skiprows = 0)
                df2 = df2.set_index("Time")
                print("Extracted file loaded successfully!!!")
                sl.sleep(2)

                times,data,ntimes = df.index.tolist(),[],0

                for time in range(len(times)):
                    # val2G,val3G = str(round(df.iloc[time,2],2)) if str(df.iloc[time,2]) != 'nan' else "", str(round(df2.iloc[time,2],2)) if time < 72 and str(df2.iloc[time,2]) != 'nan' else ""
                    val2G,val3G = str(round(df.iloc[time,2],2)) if str(df.iloc[time,2]) != 'nan' else "", str(round(df2.iloc[time,2],2)) if str(df2.iloc[time,2]) != 'nan' else ""
                    if not (val2G == '', val3G == '') == (True,True):
                        data.append([df.iloc[time,1],str(times[time]),val2G,val3G,0.0])
                        ntimes += 1
                df3 = pd.DataFrame(data,columns = ['Single Area','Visible Time','2G','3G','4G'],dtype = str)
                df3.to_excel('D:\\2020\\Python\\9Mobile\\SKYAvailability\\Output\\HourlyAvailability.xlsx','Hourly Availability', index=False)

                print(f"""Hourly Availability generated successfully!!
                {ntimes} lines of data added!""")
                sl.sleep(2)
                upload('https://103n-saapp.teleows.com/servicecreator/spl/base_dashboard/bdb_hourly_availability_import_page.spl?','HourlyAvailability.xlsx')

            except Exception as e:
                print(f'Error while generating file and uploading: {str(e)}')
                raise Exception(e)
                # pass
            
def upload(url,type):
    
    try:
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : 'D:\\2020\\Python\\9Mobile\\SKYAvailability\\Download\\'}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--headless')
        driver = webdriver.Chrome(chrome_options=chrome_options)
        driver.get(url)

        with open('D:\\2020\\Python\\9Mobile\\SKYAvailability\\System\\details.txt','r') as credentials:
            details = credentials.readlines()
            username,password = details[0].strip(),details[1].strip()

        driver.refresh()
        driver.find_element_by_id("usernameInput").send_keys(username)
        driver.find_element_by_id("password").send_keys(password)
        driver.find_element_by_id("btn_submit").click()
        sl.sleep(5)
        file = f'D:\\2020\\Python\\9Mobile\\SKYAvailability\\Output\\{type}'
        addfile = driver.find_element_by_name("_uploadFile")
        sl.sleep(2)
        addfile.send_keys(file)
        # add.send_keys(Keys.ENTER)
        sl.sleep(10)
        js = 'document.getElementsByClassName("sdm_button_primary")[0].click();'
        driver.execute_script(js)
        # driver.find_element_by_xpath('//*[@id="import_submit"]').click()
        # while driver.find_element_by_class_name('nf-el-mask'):
        #     sl.sleep(5)
        sl.sleep(15)
        driver.quit()
    except Exception as e:
        print(f'Error while uploading: {str(e)}')
        raise Exception(e)

def main():
    with open('D:\\2020\\Python\\9Mobile\\SKYAvailability\\Log\\Logs.txt','a') as log:
        try:
            deleteFiles(paths)
            if runAll('Regional Availability for GIS'):
                print('Operation completed successfully!!!')
        except Exception as e:
            print(f'An error occured at main phase: {str(e)}')
            log.write(f'{str(datetime.datetime.now())[:19]} \n')
            log.write(f'{str(e)} \n')
            log.write(f'{str(datetime.datetime.now())[:19]} \n')

if __name__ == "__main__":
    main()
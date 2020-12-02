import time 
import os
import requests, datetime
from xml.dom import minidom
from pyxlsb import open_workbook

class owsBOT:

    def __init__(self, username, password):
        self.mtnUrl = 'https://15fg-saapp.teleows.com/ws/rest/15fg/RNOC_operation_inbound/v1/rnoc_portal_tt_capture_rate'
        self.mtnUserName = username
        self.mtnPassword = password

        self.atUrl = 'https://15fg-saapp.teleows.com/ws/rest/15fg/RNOC_operation_inbound/v1/rnoc_portal_tt_capture_rate'
        self.atUserName = username
        self.atPassword = password


    def send(self,projct,randomain,siteOc,outagettc,missedttc,notappear,actuallym,num):
        project = projct
        # 'RAN-'+ws['B1']
        domain = randomain
        siteoutage_count = siteOc
        outagett_count = outagettc
        missedtt_count = missedttc
        notappear_ictom = notappear
        actually_missed = actuallym
        
        # headers = {'Content-type':'application/x-www-form-urlencoded'}
        headers = {'Content-type':'application/json'}
        
        PARAMS = {
                    "project": project,
                    "domain": domain,
                    "siteoutage_count": siteoutage_count,
                    "outagett_count": outagett_count,
                    "missedtt_count": missedtt_count,
                    "notappear_ictom": notappear_ictom,
                    "actually_missed": actually_missed,
                    "day_interval": str(num)

                }
        
        session = requests.sessions.session()

        response = session.post(
            self.mtnUrl,
            auth = (self.mtnUserName,self.mtnPassword),
            headers = headers,
            json = PARAMS,
            verify = False,
            timeout = None
        )

        if response.text == '{"result":"true"}':
            return True
        else:
            return False
    
    def MTN(self,number):
        path = "//172.16.151.77/Report Automation/MTN Report Automation/MTN TT Capture Rate/Output/MTN TT Capture/"
        
        # MTN TT Captured Rate Analysis 22-Jan-2020.xlsb
        
        file_day = datetime.date.today() - datetime.timedelta(number) #1
        dt = file_day.strftime('%d-%b-%Y')
        # print(dt)
        # if f'MTN TT Captured Rate Analysis {dt}'+'.xlsb':
        if os.path.isfile(os.path.join(path, f'MTN TT Captured Rate Analysis {dt}.xlsb')):
            filename =  f'MTN TT Captured Rate Analysis {dt}.xlsb'
            print(filename)
            
            
            wb = open_workbook(path+filename)
            ws = wb.get_sheet('PLOT')
            
            # for 2G RAN
            domain_2g = 'RAN-2G'
            siteoutage_count_2g = 0
            outagett_count_2g = 0
            missedtt_count_2g = 0
            notappear_ictom_2g = 0
            actually_missed_2g = 0
            
            # for 3G RAN
            domain_3g = 'RAN-3G'
            siteoutage_count_3g = 0
            outagett_count_3g = 0
            missedtt_count_3g = 0
            notappear_ictom_3g = 0
            actually_missed_3g = 0
            
            # for 4G RAN
            domain_4g = 'RAN-4G'
            siteoutage_count_4g = 0
            outagett_count_4g = 0
            missedtt_count_4g = 0
            notappear_ictom_4g = 0
            actually_missed_4g = 0
            
            for row in ws.rows():
                # print(row)
                for r in row:
                    # 2G FOR ALL VENDORS
                    if r.c == 1 and r.r == 2:
                        siteoutage_count_2g = r.v
                    # 
                    elif r.c == 1 and r.r == 7:
                        outagett_count_2g = r.v
                    elif r.c == 1 and r.r == 11:
                        missedtt_count_2g = r.v
                    # not appeared on ICTOM
                    elif r.c == 1 and r.r == 15:
                        notappear_ictom_2g += r.v
                    elif r.c == 2 and r.r == 15:
                        notappear_ictom_2g += r.v
                    elif r.c == 3 and r.r == 15:
                        notappear_ictom_2g += r.v
                    # actual missed
                    elif r.c == 1 and r.r == 16:
                        actually_missed_2g += r.v
                    elif r.c == 2 and r.r == 16:
                        actually_missed_2g += r.v
                    elif r.c == 3 and r.r == 16:
                        actually_missed_2g += r.v
                    
                    # 3G FOR ALL VENDORS
                    if r.c == 5 and r.r == 2:
                        siteoutage_count_3g = r.v
                    # 
                    elif r.c == 5 and r.r == 7:
                        outagett_count_3g = r.v
                    elif r.c == 5 and r.r == 11:
                        missedtt_count_3g = r.v
                    # not appeared on ICTOM
                    elif r.c == 5 and r.r == 15:
                        notappear_ictom_3g += r.v
                    elif r.c == 6 and r.r == 15:
                        notappear_ictom_3g += r.v
                    elif r.c == 7 and r.r == 15:
                        notappear_ictom_3g += r.v
                    # actual missed
                    elif r.c == 5 and r.r == 16:
                        actually_missed_3g += r.v
                    elif r.c == 6 and r.r == 16:
                        actually_missed_3g += r.v
                    elif r.c == 7 and r.r == 16:
                        actually_missed_3g += r.v
                        
                    # 4G FOR ALL VENDORS
                    if r.c == 9 and r.r == 2:
                        siteoutage_count_4g = r.v
                    # 
                    elif r.c == 9 and r.r == 7:
                        outagett_count_4g = r.v
                    elif r.c == 9 and r.r == 11:
                        missedtt_count_4g = r.v
                    # not appeared on ICTOM
                    elif r.c == 9 and r.r == 15:
                        notappear_ictom_4g += r.v
                    elif r.c == 10 and r.r == 15:
                        notappear_ictom_4g += r.v
                    elif r.c == 11 and r.r == 15:
                        notappear_ictom_4g += r.v
                    # actual missed
                    elif r.c == 9 and r.r == 16:
                        actually_missed_4g += r.v
                    elif r.c == 10 and r.r == 16:
                        actually_missed_4g += r.v
                    elif r.c == 11 and r.r == 16:
                        actually_missed_4g += r.v
                        
            print(f"""For MTN {filename}
            siteoutage_count_2g is {siteoutage_count_2g},
            outagett_count_2g is {outagett_count_2g}
            missedtt_count_2g is {missedtt_count_2g}
            notappear_ictom_2g is {notappear_ictom_2g}
            actually_missed_2g is {actually_missed_2g}
            \n""")

            print(f"""
            siteoutage_count_3g is {siteoutage_count_3g}
            outagett_count_3g is {outagett_count_3g}
            missedtt_count_3g is {missedtt_count_3g}
            notappear_ictom_3g is {notappear_ictom_3g}
            actually_missed_3g is {actually_missed_3g}
            \n""")

            print(f"""
            siteoutage_count_4g is {siteoutage_count_4g}
            outagett_count_4g is {outagett_count_4g}
            missedtt_count_4g is {missedtt_count_4g}
            notappear_ictom_4g is {notappear_ictom_4g}
            actually_missed_4g is {actually_missed_4g}
            \n""")
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying MTN 2G ... {triesCount}')
            while not (self.send('NG MTN',domain_2g,int(siteoutage_count_2g),int(outagett_count_2g),int(missedtt_count_2g),int(notappear_ictom_2g),int(actually_missed_2g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying MTN 2G again... {triesCount}')
                triesCount += 1
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying MTN 3G ... {triesCount}')
            while not (self.send('NG MTN',domain_3g,int(siteoutage_count_3g),int(outagett_count_3g),int(missedtt_count_3g),int(notappear_ictom_3g),int(actually_missed_3g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying MTN 3G again... {triesCount}')
                triesCount += 1
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying MTN 4G ... {triesCount}')
            while not (self.send('NG MTN',domain_4g,int(siteoutage_count_4g),int(outagett_count_4g),int(missedtt_count_4g),int(notappear_ictom_4g),int(actually_missed_4g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying MTN 4G again... {triesCount}')
                triesCount += 1


    def Airtel(self,number):
        path = "//172.16.151.77/Report Automation/Airtel Report Automation/Daily TT Analysis/OUTPUT/Customer/"

        # D:\Report Automation\Airtel Report Automation\Daily TT Analysis\OUTPUT\Customer

        # AIRTEL TT CAPTURE RATE ANALYSIS(INTERNAL)- Mar-16-2020

        # AIRTEL TT CAPTURE RATE ANALYSIS Mar-16-2020
        
        file_day = datetime.date.today() - datetime.timedelta(number) #1
        dt = file_day.strftime('%b-%d-%Y')
        print(dt)
        # if f'AIRTEL TT CAPTURE RATE ANALYSIS {dt}.xlsb':
        if os.path.isfile(os.path.join(path, f'AIRTEL TT CAPTURE RATE ANALYSIS {dt}.xlsb')):
            filename =  f'AIRTEL TT CAPTURE RATE ANALYSIS {dt}.xlsb'
            print(filename)
            
            wb = open_workbook(path+filename)
            ws = wb.get_sheet('PLOT')
            
            # for 2G RAN
            domain_2g = 'RAN-2G'
            siteoutage_count_2g = 0
            outagett_count_2g = 0
            missedtt_count_2g = 0
            notappear_ictom_2g = 0
            actually_missed_2g = 0
            
            # for 3G RAN
            domain_3g = 'RAN-3G'
            siteoutage_count_3g = 0
            outagett_count_3g = 0
            missedtt_count_3g = 0
            notappear_ictom_3g = 0
            actually_missed_3g = 0
            
            # for 4G RAN
            domain_4g = 'RAN-4G'
            siteoutage_count_4g = 0
            outagett_count_4g = 0
            missedtt_count_4g = 0
            notappear_ictom_4g = 0
            actually_missed_4g = 0
            
            for row in ws.rows():
                # print(row)
                for r in row:
                    # 2G FOR ALL VENDORS
                    if r.c == 5 and r.r == 5:
                        siteoutage_count_2g += r.v
                    elif r.c == 6 and r.r == 5:
                        siteoutage_count_2g += r.v
                    elif r.c == 7 and r.r == 5:
                        siteoutage_count_2g += r.v
                    # 
                    elif r.c == 5 and r.r == 7:
                        outagett_count_2g += r.v
                    elif r.c == 6 and r.r == 7:
                        outagett_count_2g += r.v
                    elif r.c == 7 and r.r == 7:
                        outagett_count_2g += r.v

                    elif r.c == 5 and r.r == 9:
                        missedtt_count_2g += r.v
                    elif r.c == 6 and r.r == 9:
                        missedtt_count_2g += r.v
                    elif r.c == 7 and r.r == 9:
                        missedtt_count_2g += r.v

                    # not appeared on ICTOM
                    elif r.c == 5 and r.r == 10:
                        notappear_ictom_2g += r.v
                    elif r.c == 6 and r.r == 10:
                        notappear_ictom_2g += r.v
                    elif r.c == 7 and r.r == 10:
                        notappear_ictom_2g += r.v

                    # actual missed
                    elif r.c == 5 and r.r == 11:
                        actually_missed_2g += r.v
                    elif r.c == 6 and r.r == 11:
                        actually_missed_2g += r.v
                    elif r.c == 7 and r.r == 11:
                        actually_missed_2g += r.v
                    
                    # 3G FOR ALL VENDORS
                    if r.c == 8 and r.r == 5:
                        siteoutage_count_3g += r.v
                    elif r.c == 9 and r.r == 5:
                        siteoutage_count_3g += r.v
                    elif r.c == 10 and r.r == 5:
                        siteoutage_count_3g += r.v
                    # 
                    elif r.c == 8 and r.r == 7:
                        outagett_count_3g += r.v
                    elif r.c == 9 and r.r == 7:
                        outagett_count_3g += r.v
                    elif r.c == 10 and r.r == 7:
                        outagett_count_3g += r.v

                    elif r.c == 8 and r.r == 9:
                        missedtt_count_3g += r.v
                    elif r.c == 9 and r.r == 9:
                        missedtt_count_3g += r.v
                    elif r.c == 10 and r.r == 9:
                        missedtt_count_3g += r.v

                    # not appeared on ICTOM
                    elif r.c == 8 and r.r == 10:
                        notappear_ictom_3g += r.v
                    elif r.c == 9 and r.r == 10:
                        notappear_ictom_3g += r.v
                    elif r.c == 10 and r.r == 10:
                        notappear_ictom_3g += r.v

                    # actual missed
                    elif r.c == 8 and r.r == 11:
                        actually_missed_3g += r.v
                    elif r.c == 9 and r.r == 11:
                        actually_missed_3g += r.v
                    elif r.c == 10 and r.r == 11:
                        actually_missed_3g += r.v

                    # 4G FOR ALL VENDORS
                    if r.c == 12 and r.r == 5:
                        siteoutage_count_4g += r.v
                    elif r.c == 13 and r.r == 5:
                        siteoutage_count_4g += r.v
                    # 
                    elif r.c == 12 and r.r == 7:
                        outagett_count_4g += r.v
                    elif r.c == 13 and r.r == 7:
                        outagett_count_4g += r.v

                    elif r.c == 12 and r.r == 9:
                        missedtt_count_4g += r.v
                    elif r.c == 13 and r.r == 9:
                        missedtt_count_4g += r.v

                    # not appeared on ICTOM
                    elif r.c == 12 and r.r == 10:
                        notappear_ictom_4g += r.v
                    elif r.c == 13 and r.r == 10:
                        notappear_ictom_4g += r.v

                    # actual missed
                    elif r.c == 12 and r.r == 11:
                        actually_missed_4g += r.v
                    elif r.c == 13 and r.r == 11:
                        actually_missed_4g += r.v
                        
                
            wss = wb.get_sheet('NOKIA')
            for row in wss.rows():
                for r in row:
                    # 2G FOR ALL VENDORS
                    if r.c == 5 and r.r == 4:
                        siteoutage_count_2g += r.v
                    # 
                    elif r.c == 5 and r.r == 6:
                        outagett_count_2g += r.v

                    elif r.c == 5 and r.r == 8:
                        missedtt_count_2g += r.v

                    # not appeared on ICTOM
                    elif r.c == 5 and r.r == 9:
                        notappear_ictom_2g += r.v

                    # actual missed
                    elif r.c == 5 and r.r == 10:
                        actually_missed_2g += r.v
                    
                    # 3G FOR ALL VENDORS
                    if r.c == 6 and r.r == 4:
                        siteoutage_count_3g += r.v
                    # 
                    elif r.c == 6 and r.r == 6:
                        outagett_count_3g += r.v

                    elif r.c == 6 and r.r == 8:
                        missedtt_count_3g += r.v

                    # not appeared on ICTOM
                    elif r.c == 6 and r.r == 9:
                        notappear_ictom_3g += r.v

                    # actual missed
                    elif r.c == 6 and r.r == 10:
                        actually_missed_3g += r.v

            print("day_interval",number)

            print(f"""For Airtel {filename}
            siteoutage_count_2g is {siteoutage_count_2g},
            outagett_count_2g is {outagett_count_2g}
            missedtt_count_2g is {missedtt_count_2g}
            notappear_ictom_2g is {notappear_ictom_2g}
            actually_missed_2g is {actually_missed_2g}
            \n""")

            print(f"""
            siteoutage_count_3g is {siteoutage_count_3g}
            outagett_count_3g is {outagett_count_3g}
            missedtt_count_3g is {missedtt_count_3g}
            notappear_ictom_3g is {notappear_ictom_3g}
            actually_missed_3g is {actually_missed_3g}
            \n""")

            print(f"""
            siteoutage_count_4g is {siteoutage_count_4g}
            outagett_count_4g is {outagett_count_4g}
            missedtt_count_4g is {missedtt_count_4g}
            notappear_ictom_4g is {notappear_ictom_4g}
            actually_missed_4g is {actually_missed_4g}
            \n""")

            triesCount = 0
            time.sleep(1)
            print(f'Trying Airtel 2G ... {triesCount}')
            while not (self.send('NG Airtel',domain_2g,int(siteoutage_count_2g),int(outagett_count_2g),int(missedtt_count_2g),int(notappear_ictom_2g),int(actually_missed_2g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying Airtel 2G again... {triesCount}')
                triesCount += 1
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying Airtel 3G ... {triesCount}')
            while not (self.send('NG Airtel',domain_3g,int(siteoutage_count_3g),int(outagett_count_3g),int(missedtt_count_3g),int(notappear_ictom_3g),int(actually_missed_3g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying Airtel 3G again... {triesCount}')
                triesCount += 1
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying Airtel 4G ... {triesCount}')
            while not (self.send('NG Airtel',domain_4g,int(siteoutage_count_4g),int(outagett_count_4g),int(missedtt_count_4g),int(notappear_ictom_4g),int(actually_missed_4g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying 9Mobile 4G again... {triesCount}')
                triesCount += 1
            
            
            
            
        
    def NMobile(self,number):
        path = "//172.16.151.77/Users/WX815001/Desktop/9TT Capture Rate/Output/"
        
        # 9MOBILE TT CAPTURE RATE FOR 8-Nov-2020.xlsb
        
        file_day = datetime.date.today() - datetime.timedelta(number) #1
        dt = file_day.strftime('%d-%b-%Y')

        if dt[0] == '0':
            dt = dt[1:]

        # print(dt)
        
        if os.path.isfile(os.path.join(path, f'9MOBILE TT CAPTURE RATE FOR {dt}.xlsb')):
            filename =  f'9MOBILE TT CAPTURE RATE FOR {dt}.xlsb'
            print(filename)
            time.sleep(2)

            wb = open_workbook(path+filename)
            ws = wb.get_sheet('PLOT')
            
            # for 2G RAN
            domain_2g = 'RAN-2G'
            siteoutage_count_2g = 0
            outagett_count_2g = 0
            missedtt_count_2g = 0
            notappear_ictom_2g = 0
            actually_missed_2g = 0
            
            # for 3G RAN
            domain_3g = 'RAN-3G'
            siteoutage_count_3g = 0
            outagett_count_3g = 0
            missedtt_count_3g = 0
            notappear_ictom_3g = 0
            actually_missed_3g = 0
            
            for row in ws.rows():
                # print(row)
                for r in row:
                    # 2G FOR ALL VENDORS
                    if r.c == 1 and r.r == 1:
                        siteoutage_count_2g = r.v
                    # Outage TT
                    elif r.c == 1 and r.r == 6:
                        outagett_count_2g += r.v
                    elif r.c == 1 and r.r == 7:
                        outagett_count_2g += r.v
                    # missing TT
                    elif r.c == 1 and r.r == 11:
                        missedtt_count_2g += r.v
                        
                    # not appeared on ICTOM
                    elif r.c == 1 and r.r == 12:
                        notappear_ictom_2g += r.v
                    # actual missed
                    elif r.c == 1 and r.r == 13:
                        actually_missed_2g += r.v
                    
                    # 3G FOR ALL VENDORS
                    if r.c == 3 and r.r == 1:
                        siteoutage_count_3g = r.v
                    # 
                    elif r.c == 3 and r.r == 6:
                        outagett_count_3g += r.v
                    elif r.c == 3 and r.r == 7:
                        outagett_count_3g += r.v
                        
                    elif r.c == 3 and r.r == 11:
                        missedtt_count_3g = r.v
                    # not appeared on ICTOM
                    elif r.c == 3 and r.r == 12:
                        notappear_ictom_3g += r.v
                    # actual missed
                    elif r.c == 3 and r.r == 13:
                        actually_missed_3g += r.v
                    
            
            
            print(f"""For 9Mobile {filename}
            siteoutage_count_2g is {siteoutage_count_2g},
            outagett_count_2g is {outagett_count_2g}
            missedtt_count_2g is {missedtt_count_2g}
            notappear_ictom_2g is {notappear_ictom_2g}
            actually_missed_2g is {actually_missed_2g}
            \n""")

            print(f"""
            siteoutage_count_3g is {siteoutage_count_3g}
            outagett_count_3g is {outagett_count_3g}
            missedtt_count_3g is {missedtt_count_3g}
            notappear_ictom_3g is {notappear_ictom_3g}
            actually_missed_3g is {actually_missed_3g}
            \n""")
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying 9Mobile 2G ... {triesCount}')
            while not (self.send('NG 9Mobile',domain_2g,int(siteoutage_count_2g),int(outagett_count_2g),int(missedtt_count_2g),int(notappear_ictom_2g),int(actually_missed_2g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying 9Mobile 2G again... {triesCount}')
                triesCount += 1
            
            triesCount = 0
            time.sleep(1)
            print(f'Trying 9Mobile 3G ... {triesCount}')
            while not (self.send('NG 9Mobile',domain_3g,int(siteoutage_count_3g),int(outagett_count_3g),int(missedtt_count_3g),int(notappear_ictom_3g),int(actually_missed_3g),number)) and (triesCount < 2):
                time.sleep(2)
                print(f'Trying 9Mobile 3G again... {triesCount}')
                triesCount += 1
        
    def AirtelTigo(self,number):
        path = "//172.16.151.77/Report Automation/AirtelTigo Report Automation/AirtelTigo TT Capture Rate/Output/"
        file_day = datetime.date.today() - datetime.timedelta(number) #1
        dt = file_day.strftime('%d-%m-%Y')
        dt = "("+dt+")"
        print(dt)
        if f'Capture Rate Report for  {dt}'+'.xlsb':
            filename =  f'Capture Rate Report for  {dt}'+'.xlsb'
            print(filename)
            
        wb = open_workbook(path+filename)
        ws = wb.get_sheet('RCA Breakdown')
        
        # for 2G RAN
        domain_2g = 'RAN-2G'
        siteoutage_count_2g = 0
        outagett_count_2g = 0
        missedtt_count_2g = 0
        notappear_ictom_2g = 0
        actually_missed_2g = 0
        
        # for 3G RAN
        domain_3g = 'RAN-3G'
        siteoutage_count_3g = 0
        outagett_count_3g = 0
        missedtt_count_3g = 0
        notappear_ictom_3g = 0
        actually_missed_3g = 0
        
        for row in ws.rows():
            # print(row)
            for r in row:
                # 2G FOR ALL VENDORS
                if r.c == 1 and r.r == 3:
                    siteoutage_count_2g = r.v
                # 
                elif r.c == 1 and r.r == 4:
                    outagett_count_2g = r.v
                    
                # 3G FOR ALL VENDORS
                if r.c == 4 and r.r == 3:
                    siteoutage_count_3g = r.v
                # 
                elif r.c == 4 and r.r == 4:
                    outagett_count_3g = r.v
                
    
        missedtt_count_2g = int(siteoutage_count_2g - outagett_count_2g)
        missedtt_count_3g = int(siteoutage_count_3g - outagett_count_3g)
        
        print(siteoutage_count_2g)
        print(outagett_count_2g)
        print(missedtt_count_2g)
        print(notappear_ictom_2g)
        print(actually_missed_2g)
        
        print(siteoutage_count_3g)
        print(outagett_count_3g)
        print(missedtt_count_3g)
        print(notappear_ictom_3g)
        print(actually_missed_3g)
        
        self.send('GH AirtelTigo',domain_2g,int(siteoutage_count_2g),int(outagett_count_2g),int(missedtt_count_2g),'','',number)
        self.send('GH AirtelTigo',domain_3g,int(siteoutage_count_3g),int(outagett_count_3g),int(missedtt_count_3g),'','',number)
        
def main():
    with open('D:\\2020\\Python\\SARAMS\\RNOCTTCaptureRate\\details.txt','r') as credentials:
        details = credentials.readlines()
        username,password = details[0].strip(),details[1].strip()

    hub = owsBOT(username,password)
    for numb in range(1,10):
        try:
            hub.MTN(numb)
        except Exception as e:
            print(e)
            pass
        # try:
        #     hub.AirtelTigo(numb)
        # except Exception as e:
        #     print(e)
        #     pass
        try:
            hub.NMobile(numb)
        except Exception as e:
            print(e)
            pass
        try:
            hub.Airtel(numb)
        except Exception as e:
            print(e)
            pass

if __name__ == "__main__":
    main()
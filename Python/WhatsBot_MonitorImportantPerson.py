import re,  os
import datetime
import time, getpass
import requests
from tkinter import *
from urllib.parse import quote_plus
from selenium import webdriver
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, ElementNotVisibleException

# from ows import *

class owsBOT:

    def __init__(self, driver, username, password):
        self.driver_ = driver
        self.cUrl = 'https://167i-sgdev.teleows.com/ws/rest/167i/RNOC_operation_inbound/UserLoginSuccessful'
        self.cUserName = username
        self.cPassword = password

        self.sUrl = 'https://167i-sgdev.teleows.com/ws/rest/167i/RNOC_operation_inbound/UserLoginSuccessful'
        self.sUserName = username
        self.sPassword = password

    def check(self):
        headers = {'Content-type':'application/x-www-form-urlencoded'}

        response = requests.get(
            self.cUrl,
            verify=False,
            auth=(self.cUserName,self.cPassword),
            params={},
            headers=headers,
            timeout=None
        )

        if response.text == '{"results":"Successfull"}':
            return True
        else:
            return False
    
    def say_hi(self):
        print("Saying hi")
    
    def forward(self):
        drv = self.driver_
        text_bubbles = drv.find_elements_by_class_name("message-in")  # message-in = receiver, message-out = sender
        tmp_queue = []

        # 'MTN Critical Incident Mgt', 'MTN Availability STF'
        try:
            for bubble in text_bubbles:
                msg_texts = bubble.find_elements_by_class_name("copyable-text")
                for msg in msg_texts:
                    #raw_msg_text = msg.find_element_by_class_name("selectable-text.invisible-space.copyable-text").text.lower()
                    # raw_msg_time = msg.find_element_by_class_name("bubble-text-meta").text        # time message sent
                    tmp_queue.append(msg.text.lower())
            if len(tmp_queue) > 0:
                group = 'MTN Critical Incident Mgt'
                # group = 'AutoBOT'
                chat = drv.find_element_by_xpath('//span[@title = "{}"]'.format(group))
                chat.click()
                time.sleep(0.5)
                drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys(tmp_queue[-1])
                drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys(Keys.ENTER) 
                chat = drv.find_element_by_xpath('//span[@title = "MTN Availability STF"]')
                chat.click()
        except StaleElementReferenceException as e:
            print(str(e))

class BotConfig(object):
    last_msg = False
    last_msg_id = False

    command_history = []
    last_command = ""

    def __init__(self, contact_list):
        self.contacts = contact_list

    def get_contacts(self):
        return self.contacts

    def set_last_chat_message(self, msg, time_id):
        self.last_msg = msg
        self.last_msg_id = time_id

    def get_last_chat_message(self):
        return self.last_msg, self.last_msg_id

    def set_last_command(self, command):
        self.last_command = command
        self.command_history.append(command)

    def get_command_history(self):
        return "You have asked the following commands: " + ", ".join(self.command_history)

class Bot(object):
    def __init__(self,smsBOT):
        self.config = BotConfig(contact_list=whatsapp_contacts())
        self.init_bot(smsBOT)

    # def init_bot(self):
    def init_bot(self,smsBOT):
        while True:
            try:
                self.poll_chat(smsBOT)

                print(f'We just might be good to go!!')
                time.sleep(10)
                # dt = datetime.datetime.now()
                # mint = dt.minute
                # sec = dt.second
                # hr = dt.hour

                # cctBOT = smsBOT(driver) 

            
                # if hr == 0 and mint == 40:
                #     cctBOT.call_sms()
                #     time.sleep(59)
                # elif hr == 0 and mint == 34:
                #     clear()
                #     time.sleep(55)
                
            except Exception as e:
                print(str(e))
                self.poll_chat(smsBOT)

    # def poll_chat(self):
    def poll_chat(self,smsBOT):
        complete_last_msg = chat_history()
        last_msg = complete_last_msg[0]

        if last_msg:
            time_id = time.strftime('%H-%M-%S', time.gmtime())

            # last_saved_msg, last_saved_msg_id = self.config.get_last_chat_message()
            last_saved_msg, last_saved_msg_id = self.config.get_last_chat_message()
            if last_saved_msg != last_msg and last_saved_msg_id != time_id:
                self.config.set_last_chat_message(msg=last_msg, time_id=time_id)
                self.config.get_last_chat_message()

                print(f'Got here, last message is +++ {last_msg}')

                is_action = is_action_message(last_msg=' '.join(last_msg.split('\n')).split(' ')[-1].lower())
                if is_action:
                    self.config.set_last_command(last_msg.split('\n')[-1].lower())
                    # self.bot_options(action=last_msg)
                    # self.bot_options(last_msg,smsBOT)
                    self.bot_options(complete_last_msg,smsBOT)

    def bot_options(self, action,smsBOT):
        # BOT = smsBOT(driver)
        BOT = smsBOT
        
        simple_menu = {                                 # function requires no extra arguments
            "hi": BOT.say_hi,
            "sms":BOT.call_sms,
        }
        simple_menu_keys = simple_menu.keys()

        try:
            # command_args = action[1:].split(" ", 1)
            # command_args = action.split('\n')[-1].split(" ", 1)
            command_args = action[0].split('\n')[-1].split(" ", 1)
            print("Command args: {cmd}".format(cmd=command_args))
            
            # if command_args[0] == 'sms':
            # if command_args[0] == '/important':
            if command_args[0].lower() == '/important':
                email_body = action
                if BOT.send_email(email_body):
                    print('HURRAY!!! Email Successfully Sent')
                else:
                    print('Village people at work!! :(')
            elif command_args[0] == 'hi':
                BOT.say_hi()
                
        except KeyError as e:
            print("Key Error Exception: {err}".format(err=str(e)))
            send_message("Wrong command. Use help to get right commands")

    @staticmethod
    def _help_commands():
        return "Commands: /hi, /help, /SMS"

class smsBOT:

    def __init__(self, driver,wait, username, password):
        self.driver_ = driver
        self.wait = wait
        self.cUserName = username
        self.cPassword = password
        self.cUrl = 'https://167i-sgdev.teleows.com/ws/rest/167i/RNOC_operation_inbound/UserLoginSuccessful'
        self.emailUrl = 'https://15fg-saapp.teleows.com/ws/rest/15fg/RNOC_operation_inbound/v1/email_send'


    def say_hi(self):
        drv = self.driver_
        print("Saying hi")
        drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys('Hi, This is your WhatsBot. You can:')
        drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys(Keys.SHIFT + Keys.ENTER)
        drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys('Input */sms* or */SMS* to get *sms*')
        drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys(Keys.SHIFT + Keys.ENTER)
        whatsapp_msg =drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]")
        time.sleep(0.5)
        whatsapp_msg.send_keys(Keys.ENTER)
    
    def call_sms(self):
        drv = self.driver_
        wait = self.wait
        print("Saying hi")
        
        group = 'Testing Bot 2'
        
        x_arg = f'//span[contains(@title,"{group}")]'
        group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg))) 
        group_title.click()

        # chat.click()
        drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]").send_keys(f"*Just Testing*")
        
        
        whatsapp_msg =drv.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]")
        time.sleep(0.5)
        whatsapp_msg.send_keys(Keys.ENTER)

    def send_email(self,email_body):
        headers = {'Content-type':'application/json'}

        # email_Content = "This message is a test sent from Whatsbot"
        email_Content = f"""
        Sender : {email_body[1]}
        <br>
        Message Time : {email_body[2]}
        <br>
        <br>
        {email_body[0].split('/Important')[0].split('/important')[0]}"""

        PARAMS = {
            "email_to":"wangmu.jerome@huawei.com; uba.kelechi.michael@huawei.com; rasheed.opeyemi1@huawei.com; godspower.obinna@huawei.com ",
            "title":"【NOTIFICATION】Whatsbot Test Email",
            "data_type":"text/html;charset=utf-8",
            "content": email_Content
            }
        
        session = requests.sessions.session()

        response = session.post(
            self.emailUrl,
            # auth = (self.cUserName,self.cPassword),
            auth = ("uwx815001","qaz!#%SSaP99"),
            headers = headers,
            json = PARAMS,
            verify = False,
            timeout = None
        )

        if response.text == '{"result":"true"}':
            return True
        else:
            return False

    def monitor(self):
        driver = self.driver_
        wait = self.wait

        group = 'Tools Automation'
        search = "_1awRl"
        
        
        x_arg = f'//div[contains(@class,"{search}")]'
        search_btn = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
        search_btn.send_keys(group)
        time.sleep(3)
        # search_btn.send_keys(Keys.ENTER)

        x_arg = f'//span[contains(@title,"{group}")]'
        group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
        group_title.click()
        # search_btn.clear()

        x_arg = f'//span[contains(@data-testid,"{"x-alt"}")]'
        close_search = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
        close_search.click()
        # search_btn.clear()

        print('Waiting for VIP message')
        # Bot()
        Bot(self)




def check(username, password):
    headers = {'Content-type':'application/x-www-form-urlencoded'}

    response = requests.get(
        'https://167i-sgdev.teleows.com/ws/rest/167i/RNOC_operation_inbound/UserLoginSuccessful',
        verify=False,
        auth=(username,password),
        params={},
        headers=headers,
        timeout=None
    )

    if response.text == '{"results":"Successfull"}':
        return True
    else:
        return False

def retrieve():
    global driver, username, password
    username = nameE.get()
    password = pwordE.get()

    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\CCT_WhatsApp\\Downloads\\'}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--ignore-certificate-errors')
    # chrome_options.add_argument('--headless')
    
    print(f'Username is {username}')
    print('Password is **********')
    
    if check(username, password):
        root.destroy()
        time.sleep(3)
        print('Username and Password correct!!!')
        time.sleep(2)
        driver = webdriver.Chrome(chrome_options=chrome_options)
        wait = WebDriverWait(driver, 600)
        driver.get('https://web.whatsapp.com')
        print("Please Scan the BarCode")
        time.sleep(5)
        owBT = smsBOT(driver,wait,username, password)
        # main(driver,wait,owBT)
        owBT.monitor()


    else:
        print('Login details not correct')
        from tkinter import messagebox
        Label(Form, text = "Username or Password incorrect", font=('arial', 8), bd=7).grid(row=3, columnspan=2)

        # messagebox.showerror("Login error", "Username or Password incorrect")

def main(driver, wait,owBT):
    
    driver.get('https://web.whatsapp.com')
    print("Please Scan the BarCode")
    time.sleep(5)
    owBT.call_sms()
    owBT.say_hi()

def chat_history():
    time.sleep(2)
    text_bubbles = driver.find_elements_by_class_name("message-in")  # message-in = receiver, message-out = sender

    # print(f'text_bubbles is {text_bubbles}')
    time.sleep(5)
    tmp_queue = []
    newLine = '\n'

    try:
        for bubble in text_bubbles:
            # if 'VIP' in [data for data in bubble.text.split(newLine)]:
            checkVIP = 'VIP' if re.search('.+VIP',[data for data in bubble.text.split(newLine)][0]) else 'Not VIP'
            if checkVIP == 'VIP':
                messageBody = bubble.text.split(newLine)
                print(f'Message possible sender is -+-+-+-+ {messageBody}')
                sender,messageTime = messageBody[0],  messageBody[-1]
                msg_texts = bubble.find_elements_by_class_name("copyable-text")
                for msg in msg_texts:
                    #raw_msg_text = msg.find_element_by_class_name("selectable-text.invisible-space.copyable-text").text.lower()
                    # raw_msg_time = msg.find_element_by_class_name("bubble-text-meta").text        # time message sent
                    # tmp_queue.append(msg.text.lower())
                    tmp_queue.append(msg.text)
            else:
                print(f'Not VIP, see ...---... {bubble.text.split(newLine)}')

        if len(tmp_queue) > 0:
            print(f'tmp_queue is {tmp_queue}')
            time.sleep(5)
            print(f'Chat History returned -=-=-=-=-= {[tmp_queue[-1],sender,messageTime]}')   # Send last message in list
            return [tmp_queue[-1],sender,messageTime]  # Send last message in list

    except StaleElementReferenceException as e:
        print(str(e))
        # Something went wrong, either keep polling until it comes back or figure out alternative

    return False


def is_action_message(last_msg):
    if last_msg == "/important":
        return True
    # if last_msg[0] == "/":
    #     return True
    # if last_msg.split('\n')[-1] == "/important":
    #     return True

    time.sleep(0.5)
    return False

def whatsapp_contacts():
    # newLine = '\n'
    contacts = driver.find_elements_by_class_name("message-in")
    contactsObject = [data.text.split('\n')[0] for data in contacts]
    print(f'contactsObject is {contactsObject}')
    time.sleep(3)
    return contactsObject

if __name__ == "__main__":
    # main()
    root = Tk()
    root.title("Whatsbot")
    root.configure(background='white')
    user = getpass.getuser()
    try:
        # root.iconbitmap(f'''C:\\Users\\{str(user)}\\Desktop\\WhatsBOT\\img\\owsBOT.ico''')
        root.iconbitmap(rf'D:\2020\Python\OBW\WhatsappMonitorImportant\System\owsBOT.ico')
    except Exception as e:
        print('Cannot find owsBOT.ico '+ str(e))
    width = 400
    height = 300
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    root.geometry("%dx%d+%d+%d" % (width, height, x, y))
    root.resizable(0, 0)

    PASSWORD = StringVar()
    USERNAME = StringVar()
    #==============================FRAMES=========================================
    Top = Frame(root, bd=2,  relief=RIDGE)

    Top.pack(side=TOP, fill=X)
    Form = Frame(root, height=200,background="white")
    Form.pack(side=TOP, pady=20)
    
    #==============================LABELS=========================================
    lbl_title = Label(Top, text = "Login with your OWS account & password", font=('arial', 13),background="#14DF14")
    lbl_title.pack(fill=X)
    lbl_username = Label(Form, text = "Username:", font=('arial', 12), bd=14, background="white")
    lbl_username.grid(row=0, sticky="e")
    lbl_password = Label(Form, text = "Password:", font=('arial', 12), bd=14, background="white")
    lbl_password.grid(row=1, sticky="w")
    lbl_text = Label(Form)
    lbl_text.grid(row=2, columnspan=2)
    
    #==============================ENTRY WIDGETS==================================
    nameE = Entry(Form, textvariable=USERNAME, font=(14),)
    nameE.grid(row=0, column=1)
    nameE.focus()
    pwordE = Entry(Form, textvariable=PASSWORD, show="*", font=(14))
    pwordE.grid(row=1, column=1)
    
    #==============================BUTTON WIDGETS=================================
    btn_login = Button(Form, text="Login", width=15, bg='#14DF14', font=('arial', 15), command=retrieve)
    btn_login.grid(pady=25, row=4, columnspan=2)
    btn_login.bind('<Return>', '')

    root.mainloop()
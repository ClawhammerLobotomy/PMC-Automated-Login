#!interpreter [optional-arg]
# -*- coding: utf-8 -*-
"""
Created for automating Plex Online login procedure for other functions to be run after
Utilizes selenium for controlling Chrome
Automatically downloads the proper Chrome webdriver required based on your Chrome install
Automatically downloads the Cumulus Plex Chrome extension
Added functionality to ignore config file and use passed in login details
Added check for failed login parameters to produce an error
6/10/2022
Added support for IAM login using legacy_login parameter.
"""
#======Begin importing======#
# Needed for referencing the main executed script
import __main__

# For selenium web driving
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from threading import Thread
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
# from pynput import keyboard

# For os file management
import os

# For compiling and checking if running compiled
import sys

# For checking for files
import pathlib
from pathlib import Path

# For configuration files
import configparser

# For reading in the PCN dictionary from an external file
import json
import csv

# For downloading chromedriver and chrome extensions
import urllib.request
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# For unzipping chromedriver and extensions
import zipfile

# For checking the file version of Chrome
from get_file_properties import *

# For creating a hotkey that will pause the running script if needed
# from pynput import keyboard

# For GUI notifications
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

import logging
from datetime import datetime
#======End package importing======#

__author__ = 'Dan Sleeman'
__copyright__ = 'Copyright 2020, PMC Automated Login'
__credits__ = ['Helmut N. https://stackoverflow.com/a/7993095']
__license__ = 'GPL-3'
__version__ = '2.3.0'
__maintainer__ = 'Dan Sleeman'
__email__ = 'sleemand@shapecorp.com'
__status__ = 'beta'

#======Set variables used for downloading the proper chromedriver version======#
chrome_browser = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'# -- ENTER YOUR Chrome.exe filepath
cb_dictionary = getFileProperties(chrome_browser) # returns whole string of version (ie. 76.0.111)
chrome_browser_version = cb_dictionary['FileVersion'].split('.')[0] # substring version to capabable version (ie. 77 / 76)
full_chrome_browser_version = cb_dictionary['FileVersion']
print(chrome_browser_version, full_chrome_browser_version)
# The key combination to check. Will create a pausefile in the directory which can be used to pause the loops.
# COMBINATIONS = [
#     {keyboard.Key.pause}
# ]

'''
Plex class
'''
class Plex:
    '''
    The main variables required to pass to the class.
    environment - accepted options are Classic and UX.
                  Determines how the program will log in
    user_id - the Plex user ID
    password - the Plex password
    company_code - the Plex company code
    pcn - Optional. PCN number that would need to be selected after login.
          Will not be needed if the account only has one PCN access or
          if using a UX login and operating in the account's main PCN
    db - Optional. Default to 'test'. Accepted values are 'test' and 'prod'.
         Can be changed via the config file after it is created.
    use_config - True/False on whether to use the config file for login details
    pcn_path - path to the pcn.json file
    legacy_login - used for IAM login process
    '''
    def __init__(self,environment, user_id, password, company_code, pcn='',
                 db='test', use_config=True, pcn_path=Path('pcn.json'),
                 legacy_login=0, cumulus=1):
        self.environment = environment.upper()
        self.user_id = user_id
        self.password = password
        self.company_code = company_code
        self.pcn = pcn
        self.db = db.upper()
        self.use_config = use_config
        self.pcn_path = pcn_path
        self.legacy_login = legacy_login
        self.cumulus = cumulus
        sql = '''Please create the pcn.json file by running the following SQL report in Plex and save it as a csv file.

SELECT
 P.Plexus_Customer_No
,P.Plexus_Customer_Name
FROM Plexus_Control_v_Customer_Group_Member P

Press OK to select the csv file.'''
        # pcn_path = Path('pcn.json')
        while not pcn_path.is_file():
            confirm = messagebox.askokcancel(title='PCN file is missing',
                                             message=sql)
            if not confirm:
                messagebox.showinfo(title='User Selected Cancel',
                                    message='The program will now close.')
                sys.exit("Process quit by user")
            self.file_path = filedialog.askopenfilename()
            if self.file_path:
                self.csv_to_json(self.file_path)
        self.pcn_dict = {}
        with open(self.pcn_path, 'r', encoding='utf-8') as pcn_config:
            self.pcn_dict = json.load(pcn_config)
        self.current = set()
        # Sets the variables for the login operation based on
        # UX or classic environment
        if self.environment == 'UX':
            self.plex_main = 'cloud.plex.com'
            self.plex_prod = ''
            self.plex_test = 'test.'
            self.plex_log_id = 'userId'
            self.plex_log_pass = 'password'
            self.plex_log_comp = 'companyCode'
            self.plex_login = 'loginButton'
            self.sso = '/sso'
        else:
            self.plex_main = '.plexonline.com'
            self.plex_prod = 'www'
            self.plex_test = 'test'
            self.plex_log_id = 'txtUserID'
            self.plex_log_pass = 'txtPassword'
            self.plex_log_comp = 'txtCompanyCode'
            self.plex_login = 'btnLogin'
            self.sso = '/signon'
        if self.legacy_login == 0:
            self.plex_log_id = 'username'
            self.plex_log_pass = 'password'
            self.plex_log_comp = 'companyCode'
            self.plex_login = 'loginButton'


    '''
    Function to take a csv file from Plex and create a
    PCN JSON file that can be used to log into specific PCNs
        if the user has multiple PCN access.
    This should only be called on initialization if
        the pcn.json file does not exist yet.
    Only required for classic logins
    '''
    def csv_to_json(self, csv_file):
        csvfile = open(csv_file, 'r', encoding='utf-8')
        jsonfile = open('pcn.json', 'w', encoding='utf-8')
        reader = csv.reader(csvfile, delimiter=',', quotechar='"')
        next(reader)
        jsonfile.write('{')
        for row in reader:
            try:
                jsonfile.write('"' + row[0] + '" : "' + row[1] + '"')
                jsonfile.write(',\n')
            except:
                continue
        jsonfile.write('"":""}')
        jsonfile.close()

    '''
    Creates the config file which can be used to change any login
        details after it is created.
    '''
    def config(self):
        if self.use_config:
    # Create the config file if it doesn't exist
            config_path = Path('config.ini')
            config = configparser.ConfigParser()
            if not config_path.is_file():
                config['Plex'] = {}
                config['Plex']['User'] = self.user_id
                config['Plex']['Pass'] = self.password
                config['Plex']['Company_Code'] = self.company_code
                config['Plex']['Database'] = self.db
                config['Plex']['PCN'] = self.pcn
                with open('config.ini', 'w+') as configfile:
                    config.write(configfile)
                config.read('config.ini')
                self.plex_db = config['Plex']['Database']
                self.plex_user = config['Plex']['User']
                self.plex_pass = config['Plex']['Pass']
                self.plex_company = config['Plex']['Company_Code']
                self.plex_pcn = config['Plex']['PCN']
            else:
                config.read('config.ini')
                self.plex_db = config['Plex']['Database']
                self.plex_user = config['Plex']['User']
                self.plex_pass = config['Plex']['Pass']
                self.plex_company = config['Plex']['Company_Code']
                self.plex_pcn = config['Plex']['PCN']
        else:
            self.plex_db = self.db
            self.plex_user = self.user_id
            self.plex_pass = self.password
            self.plex_company = self.company_code
            self.plex_pcn = self.pcn


    def token_get(self, environment):
        """
        This returns the session token for Plex
        If using UX, the whole query string is returned. 
        I.E. "__asid=######"
        """
        url = self.driver.current_url
        url_split = url.split('/')
        # url_proto = url_split[0]
        # url_domain = url_split[2]
        if environment == 'UX':
            url_token = url.split('?')[1]
            if '&' in url_token:
                url_token = url_token.split('&')[0]
            # self.url_comb = f'{url_proto}//{url_domain}'
        else:
            url_token = url_split[3]
            # self.url_comb = f'{url_proto}//{url_domain}/{url_token}'
        return url_token

    '''
    Main login function.
    This uses the config file for the variables.
    If the config file is not present,
        it will get created with the data first used in the class call.
    Logs into Plex test/prod depending on the config
    Also will select the PCN if that is configured
    '''
    def login(self, headless=0):
    # Using Chrome to access web
        extension_path = os.path.join(self.bundle_dir,'resources',
                                      'cumulus_plugin.crx')
        # print('Extension path:', extension_path)
        executable_path = os.path.join(self.bundle_dir,'resources',
                                       'chromedriver.exe')
        os.environ['webdriver.chrome.driver'] = executable_path
        # print(executable_path)
        chrome_options = Options()
        #chrome_options.add_argument("--log-level=off")
        # chrome_options.add_argument("--window-position=10,10")
        # chrome_options.add_argument("--window-size=600,800")
        chrome_options.add_argument("--log-level=3")
        if headless == 1:
            chrome_options.add_argument("--headless")
        elif self.cumulus:
            chrome_options.add_extension(extension_path)
        chrome_options.add_experimental_option("prefs",{
        "download.default_directory": f"{self.bundle_dir}\Downloads",
        "download.prompt_for_download": False,
        })
        # print("before opening webdriver.")
        self.driver = webdriver.Chrome(executable_path=executable_path,
                                  options=chrome_options)
        # print("after opening webdriver.")
    # Open the website
    # Test vs production can be configured in config file
    # Default is test
        if self.plex_db == 'PROD':
            self.plex_db = self.plex_prod
        else:
            self.plex_db = self.plex_test
        if self.legacy_login == 1:
            self.sso = ''
        self.driver.get('https://' + self.plex_db + self.plex_main + self.sso)

    # Locate id and password fields
        id_box = self.driver.find_element(By.NAME, self.plex_log_id)
        pass_box = self.driver.find_element(By.NAME, self.plex_log_pass)
        company_code = self.driver.find_element(By.NAME, self.plex_log_comp)

    # Send login information
    # This can be configured in the config file
        company_code.send_keys(self.plex_company)
        company_code.send_keys(Keys.TAB)
        id_box.send_keys(self.plex_user)
        id_box.send_keys(Keys.TAB)
        pass_box.send_keys(self.plex_pass)
        
        if (headless == 1 or self.legacy_login == 0) and self.environment == 'CLASSIC':
            main_window_handle = None
            while not main_window_handle:
                main_window_handle = self.driver.current_window_handle
            login_button = self.driver.find_element(By.ID, self.plex_login)
            login_button.click()
            signin_window_handle = None
            while not signin_window_handle:
                for handle in self.driver.window_handles:
                    if handle != main_window_handle:
                        signin_window_handle = handle
                        break
            self.driver.switch_to.window(signin_window_handle)

    # Click login
        if headless == 0 and (self.legacy_login == 1 or (self.legacy_login == 0 and self.environment == 'UX')):
            login_button = self.driver.find_element(By.ID, self.plex_login)
            login_button.click()
            print("clicked login")

    # Get URL token and store it to be used for navigation later
        url = self.driver.current_url
        url_split = url.split('/')
        url_proto = url_split[0]
        url_domain = url_split[2]
        if self.environment == 'UX':
            url_token = url.split('?')[1]
            self.url_comb = f'{url_proto}//{url_domain}'
        else:
            url_token = url_split[3]
            self.url_comb = f'{url_proto}//{url_domain}/{url_token}'

    # Click PCN
    # Default PCN can be set in the config file
    # By default this is configured to not be used
        # try:
        if self.pcn != '':
            if self.environment == 'UX':
                self.driver.get(f'{self.url_comb}/SignOn/Customer/{self.pcn}?{url_token}')
                return (self.driver, self.url_comb, url_token)
            else:
                try:
                    self.pcn = self.pcn_dict[self.plex_pcn]
                    try:
                        self.driver.find_element(By.XPATH, '//img[@alt="' + self.pcn +
                                                    '"]').click()
                    except(NoSuchElementException):
                        self.driver.find_element(By.XPATH, '//*[contains(text(), "' +
                                                    self.pcn + '")]')[0].click()
                except(IndexError):
                    raise SystemExit(self.driver, 0)
                return (self.driver, self.url_comb, "None")


    '''
    Downloads the chromedriver and cumulus plugin that will
        allow selenium to function.
    If you do not use the cumulus plugin, you will need to do extra steps to
        control Plex after login using the pop-up window
    TODO - if the login is for UX, the Cumulus plugin isn't inherintly needed.
        ind a way to separate this out?
    '''
    def download_chrome_driver(self):
        resource_path = os.path.join(self.bundle_dir,'resources')
        text_path =os.path.join(resource_path,'chromedriver.txt')
        zip_path =os.path.join(resource_path,'chromedriver.zip')
        extension_path = os.path.join(resource_path, 'cumulus_plugin.crx')
        if not os.path.exists(resource_path):
            os.mkdir(resource_path)
    # Get the latest release variable
        url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_' + chrome_browser_version
    # Connect to the URL
        urllib.request.urlretrieve(url, text_path)
        with open(text_path, 'r') as f:
            release_no = f.read()
        url = 'https://chromedriver.storage.googleapis.com/' + release_no + '/chromedriver_win32.zip'
        urllib.request.urlretrieve(url, zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(resource_path)
    # Download the Cumulus plugin
        if self.cumulus:
            extension_id = 'ohndojpkopphmijlemjgnpbbpefpaang'
            url = 'https://clients2.google.com/service/update2/crx?response=redirect&os=win&arch=x86-64&os_arch=x86-64&nacl_arch=x86-64&prod=chromecrx&prodchannel=unknown&prodversion=' + full_chrome_browser_version + '&acceptformat=crx2,crx3&x=id%3D' + extension_id + '%26uc'
            urllib.request.urlretrieve(url, extension_path)
        #return

    '''
    Checks the running script to see if it is compiled or not.
    If compiled, the resources will be stored in the temp folder
    If not, then they will be in the script's working directory
    '''
    def frozen_check(self):
        if getattr(sys, 'frozen', False):
        # Running in a bundle
            self.bundle_dir = sys._MEIPASS # pylint: disable=no-member
        else:
        # Running in a normal Python environment
            self.bundle_dir = os.path.dirname(os.path.abspath(__main__.__file__))
        return self.bundle_dir

    '''
    Pause function handling. Creates a pausefile.txt file in the working
        directory that can be used to pause the loop if the file exists.
    EX:
    Place the code snippet at the start of your loop.
        The current loop will complete after you pause the script.
    It will be paused until the file is deleted.
        # Handler to support pausing the script at the start of a loop.
        # Useful for when you need to disconnect from the network for a short time.
        # Do not pause for >90 minutes or you will be logged out
            while os.path.exists('pausefile.txt'):
                input('Remove pausefile.txt and press ENTER to continue.')
            if os.path.exists('pausefile.txt'):
                os.remove('pausefile.txt')
                print('continuing...')
            else:
                None
    '''
    # def pause(self):
    #     print ("Pausing")
    #     Path('pausefile.txt').touch()
    # def on_press(self, key):
    #     if any([key in COMBO for COMBO in COMBINATIONS]):
    #         self.current.add(key)
    #         if any(all(k in self.current for k in COMBO) for COMBO in COMBINATIONS):
    #             self.pause()
    # def on_release(self, key):
    #     if any([key in COMBO for COMBO in COMBINATIONS]):
    #         self.current.remove(key)
    # def listen(self):
    #     with keyboard.Listener(on_press=self.on_press,
    #                            on_release=self.on_release) as listener:
    #         listener.join()
    def ux_click_button(self, name, search=0):
        """
        Takes the text from a UX button and searches for it.
        When found, triggers the mousedown function.
        """
        if self.environment == "CLASSIC":
            print("This only works in UX.")
            return
        self.name = name
        self.search = search
        self.script = """
        var buttons = document.getElementsByTagName('button')
        var searchText = "{name}"
        var found

        for (var i =0;i<buttons.length;i++){{
        if(buttons[i].textContent == searchText){{
            found = buttons[i]
            if ({search}==1){{
            $(found).trigger("mousedown")}}
            else {{
                found.click()
            }}
        break}}}}
        """.format(name=self.name, search=self.search)
        self.driver.execute_script(self.script)


    def edi_upload(self, file_list):
        """
        Takes a list of files and uploads them to the EDI log.
        
        Should be logged in to the PCN first before calling.
        """
        if self.environment == 'UX':
            print("This function only works in classic.")
            return
        self.driver.get(f'{self.url_comb}/EDI/EDI_Log2.asp')
        for f in file_list:
            self.driver.get(f'{self.url_comb}/Upload_Apps/'
                            f'EDI_Upload_Process.asp')
            input_box = self.driver.find_element(By.ID, 'FILE1')
            input_box.send_keys(f)
            # time.sleep(100000)
            timeout = 50
            try:
                element_present = EC.presence_of_element_located((By.ID,
                                                        'flttxtCustomer_No'))
                WebDriverWait(self.driver, timeout).until(element_present)
            except TimeoutException:
                continue


    def switch_pcn(self, pcn):
        """
        Switches pcns.
        """
        self.pcn = str(pcn)
        if self.environment == 'UX':
            url_token = self.token_get('UX')
            self.driver.get(f'{self.url_comb}/SignOn/Customer/{self.pcn}?{url_token}')
        else:
            url_token = self.token_get('Classic')
            self.driver.get(f'{self.url_comb}/Modules/SystemAdministration'
                                   f'/MenuSystem/MenuCustomer.aspx')
            # Click on the PCN link
            pcn = self.pcn_dict[self.pcn]
            try:
                self.driver.find_element(By.XPATH, '//img[@alt="' + pcn
                                                + '"]').click()
            except NoSuchElementException:
                self.driver.find_elements(By.XPATH, '//*[contains(text(), "'
                                            + pcn + '")]')[0].click()

    def setup_logger(self, name, log_file='log.log', file_format='DAILY',
                     level=logging.DEBUG, formatter=None):
        """
        To setup as many loggers as you want.

        The log file name will have the date pre-pended to whatever is added as the
        log file name

        Default formatter %(asctime)s - %(name)s - %(levelname)s - %(message)s
        """
        if formatter == None:
            self.formatter = '%(asctime)s - %(name)s - %(levelname)s - '+\
                             '%(message)s'
        else:
            self.formatter = formatter
        self.name = str(name)
        self.file_format = str(file_format).upper()
        self.today = datetime.now()
        self.formatter = logging.Formatter(self.formatter)
        if self.file_format == 'DAILY':
            self.log_date = self.today.strftime("%Y_%m_%d_")
        elif self.file_format == 'MONTHLY':
            self.log_date = self.today.strftime("%Y_%m_")
        else:
            self.log_date = ''
        self.log_file = self.log_date + log_file
        self.handler = logging.FileHandler(self.log_file, mode='a', encoding='utf-8')        
        self.handler.setFormatter(self.formatter)

        self.logger = logging.getLogger(self.name)
        self.logger.setLevel(level)
        self.logger.addHandler(self.handler)

        return self.logger


    def make_csv_dict(self, row):
        """
        I find this is a very useful way to get csv column header references
        Using the names of the columns works even when the position of them
          changes.
        """
        self.row = row
        cd = {}
        for x, i in enumerate(row):
            if i in cd.keys():
                i = i+'_'+str(x)
            cd[i] = x
        return cd


    def rest_create_json(self, *args):
        # print(args)
        return {'inputs': dict(args)}


    def rest_api_query(self, api_id, query, db, authentication):
        self.api_id = api_id
        self.query = query
        self.authentication = authentication
        if db.upper() == 'TEST':
            print("Using test database")
            self.db = 'test.'
        else:
            self.db = ''
        self.url = f'https://{self.db}cloud.plex.com/api/datasources/{self.api_id}/execute'
        self.session = requests.Session()
        self.retry = Retry(connect=3, backoff_factor = 0.5)
        self.adapter = HTTPAdapter(max_retries = self.retry)
        self.session.mount('https://', self.adapter)
        self.response = self.session.post(self.url, json=self.query, auth=self.authentication)
        # response = requests.post(url, json=query, auth=authentication)
        print(self.response)
        self.json_data = json.loads(self.response.text)
        # print(json_data)
        return self.json_data

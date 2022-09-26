from telnetlib import EC
from webbrowser import Chrome

from selenium import webdriver
from threading import Thread
import xlrd
import xlwt
import math
import os
import zipfile

from selenium.webdriver.support.wait import WebDriverWait
from xlutils.copy import copy
from xlwt import Workbook
import base64
import time
import datetime
import pandas as pd
from selenium.webdriver.common.by import By
from random import randrange
# from selenium.common.exceptions import NoSuchElementException
# from python_anticaptcha import AnticaptchaClient, ImageToTextTask
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import random
from queue import Queue
from io import BytesIO
import slack
from fake_useragent import UserAgent
from selenium.webdriver.chrome.options import Options
from anticaptchaofficial.recaptchav3proxyless import *
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver.v2 as uc
import chromedriver_binary
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import pyautogui as pyautogui


slack_client = slack.WebClient(token='xoxb-2676006450832-2652366448387-He07Si8e1kkte8NrQrDSauUA')
url = "https://ais.usvisa-info.com/en-ae/niv/users/sign_in"
user_name = "jilda1987"
password = "c357e5-3455c2-619eae-3a0683-207b8a"

def page_one():
    proxy = assign_random_proxy()
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--no-first-run")
    options.add_argument('--proxy-server=%s' % proxy)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
    driver.get(url)
    time.sleep(5)
    pyautogui.typewrite(user_name)
    pyautogui.press('tab')
    pyautogui.typewrite(password)
    pyautogui.press('enter')

    time.sleep(3)
    return driver


def login_page(data,driver):
    time.sleep(3)
    driver.find_element(by=By.ID,value='user_email').send_keys(form_data['email'])
    time.sleep(2)
    driver.find_element(by=By.ID,value='user_password').send_keys(form_data['password'])
    button = driver.find_element(by=By.ID,value='policy_confirmed')
    driver.execute_script(
        "document.getElementsByClassName('icheck-area-20')[0].click()"
        #     "document.getElementById('g-recaptcha-response').value='sdfghjk';"
    )
    time.sleep(2)
    print(button.get_attribute('innerHTML'))
    time.sleep(2)
    # button.click()
    driver.find_element(by=By.NAME,value='commit').click()
    time.sleep(3)
    continues = driver.find_elements(by=By.XPATH,value="//a[contains(text(), 'Continue')]")
    print(len(continues))
    ivrNumbers = driver.find_elements(by=By.XPATH,value="//div[@class='medium-12 columns text-right']")
    ivrNumber = ivrNumbers[0].text.split(': ')
    continues[0].click()
    # print(driver.find_element(by=By.TAG_NAME,value="div").find_element(by=By.CLASS_NAME,value="medium-12").text)
    # driver.find_element(by=By.CLASS_NAME, value="application attend_appointment card success") \
    #     .find_element(by=By.CLASS_NAME, value="button primary small") \
    #     .click()

    Schedule_page(driver,data)


    # groups_page(driver,data)



def Schedule_page(driver,data):
    try:
        driver.find_element(by=By.XPATH,value="//li[@class='accordion-item'][1]").click()
        time.sleep(1)
        driver.find_element(by=By.XPATH,value="//a[contains(text(), 'Schedule Appointment')]").click()
    except:

        driver.find_element(by=By.XPATH,value="//li[@class='accordion-item'][2]").click()
        time.sleep(1)
        driver.find_element(by=By.XPATH,value="//a[contains(text(), 'Reschedule Appointment')]").click()
    time.sleep(3)
    try:
        driver.find_element(by=By.ID,value='appointments_consulate_appointment_facility_id')\
            .find_element(by=By.XPATH,value="option[text()='Dubai']").click()
    except:
        driver.find_element_by_xpath("//input[@class='button primary']").click()
        time.sleep(6)

    expected_month = data.get('month').split("-")
    find_date = 0
    state_code = '96'
    first = 0
    session_close = 0
    while not find_date:
        irand = randrange(4, 7)
        time.sleep(irand)
        first = 1

        driver.find_element(by=By.ID,value='appointments_consulate_appointment_date').click()

        time.sleep(1)

        time.sleep(1)

        available = 0

        index = 0
        session_close = 0
        while not available:
            month_name = driver.execute_script("""
                                   var firstDate = document.querySelectorAll('[data-handler="selectDay"]')[0]
                                   var month_name = '';
                                   if (typeof(firstDate) !== 'undefined'){
                                   console.log('get available')
                                   month_name = document.querySelectorAll('[data-handler="selectDay"]')[0].parentElement.parentElement.parentElement.previousSibling.childNodes[1].childNodes[0].innerHTML;
                                   return month_name
                                   }
                                   """)
            print(month_name)
            if month_name is None:
                driver.find_element(by=By.XPATH,value="//a[@title='Next']").click()
                time.sleep(3)
            else:
                print(expected_month)
                if month_name in expected_month:
                    print("This is month: {}".format(month_name))
                    available = 1
                    find_date = 1
                    date_confirmed = 1
                else:
                    available = 1
                    try:
                        driver.find_element(by=By.CLASS_NAME,value="stepPending").click()
                    except:
                        driver.find_element(by=By.CLASS_NAME,value="stepPending").click()



            index = index + 1
        if state_code == '96':
            state_code = '97'
        else:
            state_code = '96'

    if find_date:
        driver.find_element(by=By.XPATH,value="//td[@class=' undefined']").click()
        driver.find_element(by=By.ID,value="appointments_consulate_appointment_time").click()
        try:
            driver.find_element(by=By.XPATH,value="//select[@id='appointments_consulate_appointment_time']/option[2]").click()
        except:
            driver.find_element(by=By.ID,value="appointments_consulate_appointment_time").click()
            time.sleep(2)
            driver.find_element(by=By.XPATH,value=
                "//select[@id='appointments_consulate_appointment_time']/option[1]").click()
    driver.find_element(by=By.ID,value='appointments_submit').click()


def assign_random_proxy():
    threads = []
    allProxies = []
    baseDNS = "premium.residential.proxyrack.net:"

    portRange = 5
    basePort = 10000
    for i in range(1, portRange):
        allProxies.append(baseDNS + str(basePort))
        basePort += 1
    return random.choice(allProxies)

from selenium.webdriver.common.keys import Keys
def checkProxy(prox):
    user_name = "jilda1987"
    password = "c357e5-3455c2-619eae-3a0683-207b8a"
    # p = prox['IP Address'] + ':' + prox['Port']
    print(prox," -> ","working")
    options = uc.ChromeOptions()
    p = "premium.residential.proxyrack.net:10000"
    options.add_argument('--proxy-server=%s' % prox)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # try:
    driver.get('https://ipinfo.io/')
    time.sleep(5)
    pyautogui.typewrite(user_name)
    pyautogui.press('tab')
    pyautogui.typewrite(password)
    pyautogui.press('enter')
    # alert = driver.switch_to.alert()
    # alert.send_keys(user_name)
    # alert.send_keys(Keys.TAB)
    # alert.send_keys(password)
    # alert.accept()
    # except Exception as e:
    #     print("Error {}".format(e))


    # except:
    #     print("{} ip is not working".format(p))





if __name__ == '__main__':
    loc = ("data-isl.xls")

    rb = xlrd.open_workbook(loc)
    sheet = rb.sheet_by_index(0)
    # For row 0 and column 0
    sheet.cell_value(0, 0)
    print(sheet.nrows)
    if sheet.cell_value(2, 0) != '' and sheet.cell_value(2, 4) == '0':
        form_data = {}
        form_data["email"] = sheet.cell_value(2, 0)
        form_data["password"] = sheet.cell_value(2, 1)
        # form_data["center"] = sheet.cell_value(n, 2)
        form_data["month"] = sheet.cell_value(2, 3)
        form_data["status"] = sheet.cell_value(2, 4)
        form_data["cell_no"] = 2

        form_data["ip"] = sheet.cell_value(2, 5)
        form_data["port"] = sheet.cell_value(2, 6)
        form_data["ip_user"] = sheet.cell_value(2, 7)
        form_data["ip_pass"] = sheet.cell_value(2, 8)
        que = Queue()
        user_index = 0
    driver = page_one()
    print(form_data)
    login_page(form_data,driver)

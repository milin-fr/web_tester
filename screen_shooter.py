import simplejson
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import random
import sys
import tkinter
from tkinter import Label
from tkinter import Button
from tkinter import Entry
from tkinter import OptionMenu
from tkinter import StringVar
from tkinter import Listbox
from tkinter import Scrollbar
from tkinter import messagebox
from tkinter import N, S, E, W
from tkinter import VERTICAL
from threading import Thread
from tkinter.messagebox import showinfo

'''Below module allows us to interact with Windows files.'''
import os

'''below 3 lines allows script to check the directory where it is executed, so it knows where to crete the excel file. I copied the whole block from stack overflow'''
abspath = os.path.abspath(__file__)
current_directory = os.path.dirname(abspath)
os.chdir(current_directory)



options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36")

page_to_screen = "https://eu.shop.battle.net/"

driver = webdriver.Chrome(options=options)

driver.set_window_size(1920, 1080)

languages = ["en-gb", "de-de", "fr-fr", "es-es", "it-it", "pl-pl", "pt-br", "ru-ru"]

def take_screenshots():
    driver.get(page_to_screen + languages[0])
    time.sleep(5)
    driver.find_element_by_xpath("""/html/body/storefront-root/storefront-home-page/div/main/storefront-family-bar/div/div/div[1]/div/storefront-link[1]/a""").click()
    time.sleep(6)
    driver.find_element_by_xpath("""//*[@id="group-link-services"]/span""").click()
    time.sleep(7)
    for language in languages:
        driver.get(page_to_screen + language)
        time.sleep(2)
        driver.save_screenshot("landing_page_" + language + ".png")
        time.sleep(1)
        driver.find_element_by_xpath("""/html/body/storefront-root/storefront-home-page/div/main/storefront-family-bar/div/div/div[1]/div/storefront-link[1]/a""").click()
        time.sleep(1)
        driver.find_element_by_xpath("""//*[@id="group-link-services"]/span""").click()
        time.sleep(1)
        driver.save_screenshot("wow_services_" + language + ".png")
    showinfo("Done","Screenshots saved!")

take_screenshots()

driver.quit()
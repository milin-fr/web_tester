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
from pywinauto.keyboard import SendKeys

'''Below module allows us to interact with Windows files.'''
import os

'''below 3 lines allows script to check the directory where it is executed, so it knows where to crete the excel file. I copied the whole block from stack overflow'''
abspath = os.path.abspath(__file__)
current_directory = os.path.dirname(abspath)
os.chdir(current_directory)

row_counter = 0
dictionary_of_action_selector_per_row = {}
dictionary_of_action_selector_variable_per_row = {}

dictionary_of_action_input_per_row = {}

list_of_actions = ["Go to", "Click element", "Scroll to", "Enter text", "Take screenshot", "Wait"]


def remove_action_row():
    #destroys all GUI elements (game selecetion and time selection buttons) created for given row
    global row_counter
    if row_counter > 0:
        row_counter -= 1
        dictionary_of_action_selector_per_row[row_counter].destroy()
        dictionary_of_action_input_per_row[row_counter].destroy()
        del dictionary_of_action_selector_variable_per_row[row_counter]


def add_action_row():
    global row_counter
    global list_of_actions
    dictionary_of_action_selector_variable_per_row[row_counter] = StringVar()
    dictionary_of_action_selector_variable_per_row[row_counter].set(list_of_actions[0])
    dictionary_of_action_selector_per_row[row_counter] = OptionMenu(main_window_of_gui, dictionary_of_action_selector_variable_per_row[row_counter], *list_of_actions)
    dictionary_of_action_selector_per_row[row_counter].config(width=13)
    dictionary_of_action_selector_per_row[row_counter].grid(row = row_counter + 1, column = 0)

    dictionary_of_action_input_per_row[row_counter] = Entry(main_window_of_gui)
    dictionary_of_action_input_per_row[row_counter].config(width=55)
    dictionary_of_action_input_per_row[row_counter].grid(row = row_counter + 1, column = 1)
    row_counter += 1

def go_to(driver, page_link):
    driver.get(page_link)

def click_element(driver, element_to_click):
    driver.find_element_by_xpath(element_to_click).click()

def scroll_to(driver, element_to_scroll_to):
    actions = ActionChains(driver)
    element = driver.find_element_by_xpath(element_to_scroll_to)
    actions.move_to_element(element).perform()

def enter_text(input_text):
    SendKeys(input_text, pause = 0.1)

def create_browser():
    options = webdriver.ChromeOptions()
    if 1 == 0:
        options.add_argument('headless')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1920, 1080)
    return driver

def take_screenshot(driver, screenshot_name):
    driver.save_screenshot(screenshot_name)

def single_action(driver, line_number, language):
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Go to":
        page_link = dictionary_of_action_input_per_row[line_number].get()
        page_link = page_link.lower()
        page_link = page_link.replace("/en-gb", language)
        go_to(driver, page_link)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Click element":
        element_to_click = dictionary_of_action_input_per_row[line_number].get()
        click_element(driver, element_to_click)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Scroll to":
        element_to_scroll_to = dictionary_of_action_input_per_row[line_number].get()
        scroll_to(driver, element_to_scroll_to)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Enter text":
        input_text = dictionary_of_action_input_per_row[line_number].get()
        enter_text(input_text)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Take screenshot":
        screenshot_name = dictionary_of_action_input_per_row[line_number].get() + "_" + language[1:] + ".png"
        take_screenshot(driver, screenshot_name)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Wait":
        time_to_wait = eval(dictionary_of_action_input_per_row[line_number].get())
        try:
            time.sleep(time_to_wait)
        except:
            time.sleep(1)

languages = ["/en-gb", "/de-de", "/fr-fr", "/es-es", "/it-it", "/pl-pl", "/pt-br", "/ru-ru"]

def perform_actions():
    driver = create_browser()
    for language in languages:
        for line_number in range(row_counter):
            single_action(driver, line_number, language)
    driver.close()



#page_to_screen = "https://eu.shop.battle.net/"

#driver = webdriver.Chrome(options=options)

#driver.set_window_size(1920, 1080)



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


main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("Screen-shooter WIP")
main_window_of_gui.wm_attributes("-topmost", 1)

label_empty_space = Label(main_window_of_gui ,text = "", width = 65)
label_empty_space.grid(row = 0, column = 0, columnspan = 2)
button_add_action_row = Button(main_window_of_gui, text = "Add row", command = add_action_row)
button_add_action_row.grid(row = 0, column = 5)
button_remove_action_row = Button(main_window_of_gui, text = "Remove row", command = remove_action_row)
button_remove_action_row.grid(row = 0, column = 6)
button_perform_actions = Button(main_window_of_gui, text = "Perform actions", height = 3, command = perform_actions)
button_perform_actions.grid(row = 1, column = 5, rowspan = 3, columnspan = 2)

main_window_of_gui.mainloop()

#driver.quit()
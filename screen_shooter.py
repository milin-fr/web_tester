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
from tkinter import Label, Button, Entry, Checkbutton, OptionMenu
from tkinter import StringVar, IntVar
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

#languages = ["/en-gb", "/de-de", "/fr-fr", "/es-es", "/it-it", "/pl-pl", "/pt-br", "/ru-ru"]


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
    dictionary_of_action_selector_per_row[row_counter].grid(row = row_counter + 1, column = 0, columnspan = 2)

    dictionary_of_action_input_per_row[row_counter] = Entry(main_window_of_gui)
    dictionary_of_action_input_per_row[row_counter].config(width=55)
    dictionary_of_action_input_per_row[row_counter].grid(row = row_counter + 1, column = 2, columnspan = 7)
    row_counter += 1

def go_to(driver, line_number, language):
    list_of_language_tags = ["/en-us", "/en-gb", "/de-de", "/fr-fr", "/es-es", "/it-it", "/pl-pl", "/pt-br", "/ru-ru"]
    page_link = dictionary_of_action_input_per_row[line_number].get()
    page_link = page_link.lower()
    for tag in list_of_language_tags:
        if tag in page_link:
            page_link = page_link.replace(tag, language)
    insert_text("Opening the " + page_link)
    driver.get(page_link)
    time.sleep(1)

def click_element(driver, line_number):
    time.sleep(3)
    element_to_click = dictionary_of_action_input_per_row[line_number].get()
    insert_text("Clicking element with Xpath: " + element_to_click)
    driver.find_element_by_xpath(element_to_click).click()

def scroll_to(driver, line_number):
    #actions = ActionChains(driver)
    time.sleep(3)
    element_to_scroll_to = dictionary_of_action_input_per_row[line_number].get()
    insert_text("Scrolling to element with Xpath: " + element_to_scroll_to)
    element = driver.find_element_by_xpath(element_to_scroll_to)
    driver.execute_script("arguments[0].scrollIntoView();", element)

def get_last_clicked_xpath():
    last_clicked_xpath = ""
    for i in range(row_counter):
        if dictionary_of_action_selector_variable_per_row[i].get() == "Click element":
            last_clicked_xpath = dictionary_of_action_input_per_row[i].get()
    return last_clicked_xpath

def enter_text(driver, line_number):
    last_clicked_xpath = get_last_clicked_xpath()
    input_text = dictionary_of_action_input_per_row[line_number].get()
    insert_text("Typing " + '"' + input_text + '"' + " in element " + '"' + last_clicked_xpath + '"')
    driver.find_element_by_xpath(last_clicked_xpath).send_keys(input_text)

def create_browser():
    options = webdriver.ChromeOptions()
    if display_browser_var.get() == 0:
        options.add_argument('headless')
        insert_text("Browser is in hidden mode.")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1920, 1080)
    return driver

def take_screenshot(driver, line_number, language):
    screenshot_name = dictionary_of_action_input_per_row[line_number].get() + "_" + language[1:] + ".png"
    insert_text("Taking screenshot: " + screenshot_name)
    driver.save_screenshot(screenshot_name)

def single_action(driver, line_number, language):
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Go to":
        go_to(driver, line_number, language)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Click element":
        click_element(driver, line_number)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Scroll to":
        scroll_to(driver, line_number)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Enter text":
        enter_text(driver, line_number)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Take screenshot":
        take_screenshot(driver, line_number, language)
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Wait":
        time_to_wait = eval(dictionary_of_action_input_per_row[line_number].get())
        try:
            insert_text("Waiting " + str(time_to_wait) + " second(s).")
            time.sleep(time_to_wait)
        except:
            insert_text("Waiting 1 second.")
            time.sleep(1)

def get_the_list_of_languages():
    list_of_languages = []
    if language_en_gb_var.get() == 1:
        list_of_languages.append("/en-gb")
    if language_de_de_var.get() == 1:
        list_of_languages.append("/de-de")
    if language_fr_fr_var.get() == 1:
        list_of_languages.append("/fr-fr")
    if language_es_es_var.get() == 1:
        list_of_languages.append("/es-es")
    if language_it_it_var.get() == 1:
        list_of_languages.append("/it-it")
    if language_pt_br_var.get() == 1:
        list_of_languages.append("/pt-br")
    if language_pl_pl_var.get() == 1:
        list_of_languages.append("/pl-pl")
    if language_ru_ru_var.get() == 1:
        list_of_languages.append("/ru-ru")
    return list_of_languages

def perform_actions():
    languages = get_the_list_of_languages()
    if len(languages) == 0:
        showinfo("Warning!", "Please, select at least one language")
    else:
        def slow_magic():
            insert_text("Starting the process!")
            insert_text("Please, don't interact with interface until done. You can minimize this window.")
            driver = create_browser()
            for language in languages:
                for line_number in range(row_counter):
                    insert_text("Performing step " + str(line_number + 1) + " for " + language[1:])
                    single_action(driver, line_number, language)
            driver.close()
            showinfo("Done!", "You can check out the results.")
        executing = Thread(target=slow_magic)
        executing.start()

def toggle_all_languages():
    if language_all_var.get() == 1:
        language_en_gb_var.set(1)
        language_de_de_var.set(1)
        language_fr_fr_var.set(1)
        language_es_es_var.set(1)
        language_it_it_var.set(1)
        language_pt_br_var.set(1)
        language_pl_pl_var.set(1)
        language_ru_ru_var.set(1)
    if language_all_var.get() == 0:
        language_en_gb_var.set(0)
        language_de_de_var.set(0)
        language_fr_fr_var.set(0)
        language_es_es_var.set(0)
        language_it_it_var.set(0)
        language_pt_br_var.set(0)
        language_pl_pl_var.set(0)
        language_ru_ru_var.set(0)

def insert_text(text):
    text_box.insert('end', text)
    text_box.see("end")

main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("Screen-shooter WIP")
main_window_of_gui.wm_attributes("-topmost", 1)

language_all_var = IntVar()
language_en_gb_var = IntVar()
language_de_de_var = IntVar()
language_fr_fr_var = IntVar()
language_es_es_var = IntVar()
language_it_it_var = IntVar()
language_pt_br_var = IntVar()
language_pl_pl_var = IntVar()
language_ru_ru_var = IntVar()

display_browser_var = IntVar()

#label_empty_space = Label(main_window_of_gui ,text = "", width = 65)
#label_empty_space.grid(row = 0, column = 0, columnspan = 2)
language_all_toggle = Checkbutton(main_window_of_gui, text="All languages", variable=language_all_var, command = toggle_all_languages)
language_all_toggle.grid(row = 0, column = 8)
language_en_gb_toggle = Checkbutton(main_window_of_gui, text="en-GB", variable=language_en_gb_var)
language_en_gb_toggle.grid(row = 0, column = 0)
language_de_de_toggle = Checkbutton(main_window_of_gui, text="de_DE", variable=language_de_de_var)
language_de_de_toggle.grid(row = 0, column = 1)
language_fr_fr_toggle = Checkbutton(main_window_of_gui, text="fr-FR", variable=language_fr_fr_var)
language_fr_fr_toggle.grid(row = 0, column = 2)
language_es_es_toggle = Checkbutton(main_window_of_gui, text="es-ES", variable=language_es_es_var)
language_es_es_toggle.grid(row = 0, column = 3)
language_it_it_toggle = Checkbutton(main_window_of_gui, text="it-IT", variable=language_it_it_var)
language_it_it_toggle.grid(row = 0, column = 4)
language_pt_br_toggle = Checkbutton(main_window_of_gui, text="pt-BR", variable=language_pt_br_var)
language_pt_br_toggle.grid(row = 0, column = 5)
language_pl_pl_toggle = Checkbutton(main_window_of_gui, text="pl-PL", variable=language_pl_pl_var)
language_pl_pl_toggle.grid(row = 0, column = 6)
language_ru_ru_toggle = Checkbutton(main_window_of_gui, text="ru-RU", variable=language_ru_ru_var)
language_ru_ru_toggle.grid(row = 0, column = 7)

display_browser_toggle = Checkbutton(main_window_of_gui, text="Display browser", variable=display_browser_var)
display_browser_toggle.grid(row = 0, column = 10)

button_add_action_row = Button(main_window_of_gui, text = "Add row", command = add_action_row)
button_add_action_row.grid(row = 1, column = 9)
button_remove_action_row = Button(main_window_of_gui, text = "Remove row", command = remove_action_row)
button_remove_action_row.grid(row = 1, column = 10)
button_perform_actions = Button(main_window_of_gui, text = "Perform actions", height = 3, command = perform_actions)
button_perform_actions.grid(row = 2, column = 9, rowspan = 3, columnspan = 2)

text_box = Listbox(main_window_of_gui, height=8)
text_box.grid(column=0, row=100, columnspan=12, sticky=(N,W,E,S))  # columnspan − How many columns widgetoccupies; default 1.
main_window_of_gui.grid_columnconfigure(0, weight=1)
main_window_of_gui.grid_rowconfigure(12, weight=1)
#scroll bar
my_scrollbar = Scrollbar(main_window_of_gui, orient=VERTICAL, command=text_box.yview)
my_scrollbar.grid(column=12, row=100, sticky=(N,S))
#attaching scroll bar to text box
text_box['yscrollcommand'] = my_scrollbar.set


main_window_of_gui.mainloop()

#driver.quit()
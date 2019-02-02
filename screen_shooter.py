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
from tkinter import Toplevel
from tkinter.filedialog import askopenfilename
from threading import Thread
from tkinter.messagebox import showinfo
#from pywinauto.keyboard import SendKeys

'''Below module allows us to interact with Windows files.'''
import os

'''below 3 lines allows script to check the directory where it is executed, so it knows where to crete the excel file. I copied the whole block from stack overflow'''
abspath = os.path.abspath(__file__)
current_directory = os.path.dirname(abspath)
os.chdir(current_directory)

'''Below module allows us to interact with Excel files'''
import openpyxl
from openpyxl import Workbook, load_workbook

row_counter = 0
dictionary_of_action_index_per_row = {}
dictionary_of_action_selector_per_row = {}
dictionary_of_action_selector_variable_per_row = {}
dictionary_of_action_input_per_row = {}

dictionary_of_sequences = {}

list_of_actions = ["Open (URL)", "Click (Xpath)", "Send (text)", "Screenshot (save as)", "Wait (seconds)", "Open (URL) +", "Click (Xpath) +", "Send (text) +", "Screenshot (save as) +"]

#languages = ["/en-gb", "/de-de", "/fr-fr", "/es-es", "/it-it", "/pl-pl", "/pt-br", "/ru-ru"]

def build_list_of_actions():
    global row_counter
    list_of_complex_imputs = ["Open (URL) +", "Click (Xpath) +", "Send (text) +", "Screenshot (save as) +", "Language +", "Browser +"]



def remove_action_row():
    #destroys all GUI elements (game selecetion and time selection buttons) created for given row
    global row_counter
    if row_counter > 0:
        row_counter -= 1
        dictionary_of_action_index_per_row[row_counter].destroy()
        dictionary_of_action_selector_per_row[row_counter].destroy()
        dictionary_of_action_input_per_row[row_counter].destroy()
        del dictionary_of_action_selector_variable_per_row[row_counter]

def add_action_row():
    global row_counter
    global list_of_actions
    dictionary_of_action_index_per_row[row_counter] = Label(main_window_of_gui, text = str(row_counter + 1))
    dictionary_of_action_index_per_row[row_counter].grid(row = row_counter + 1, column = 0)
    
    dictionary_of_action_selector_variable_per_row[row_counter] = StringVar()
    dictionary_of_action_selector_variable_per_row[row_counter].set(list_of_actions[0])
    dictionary_of_action_selector_per_row[row_counter] = OptionMenu(main_window_of_gui, dictionary_of_action_selector_variable_per_row[row_counter], *list_of_actions)
    dictionary_of_action_selector_per_row[row_counter].config(width=18)
    dictionary_of_action_selector_per_row[row_counter].grid(row = row_counter + 1, column = 1, columnspan = 3)

    dictionary_of_action_input_per_row[row_counter] = Entry(main_window_of_gui)
    dictionary_of_action_input_per_row[row_counter].config(width=55)
    dictionary_of_action_input_per_row[row_counter].grid(row = row_counter + 1, column = 4, columnspan = 6)

    row_counter += 1

def close_cookie_disclaimer(driver):
    driver.find_element_by_xpath('''//*[@id="cookie-compliance-agree"]''').click()

def go_to(driver, line_number, language):
    page_link = dictionary_of_action_input_per_row[line_number].get()
    if language != " default url":
        list_of_language_tags = ["/en-us", "/en-gb", "/de-de", "/fr-fr", "/es-es", "/it-it", "/pl-pl", "/pt-br", "/ru-ru"]
        page_link = page_link.lower()
        for tag in list_of_language_tags:
            if tag in page_link:
                page_link = page_link.replace(tag, language)
    insert_text("Opening the " + page_link)
    driver.get(page_link)
    try:
        close_cookie_disclaimer(driver)
    except:
        pass

def click_element(driver, line_number):
    element_to_click = dictionary_of_action_input_per_row[line_number].get()
    insert_text("Clicking element with Xpath: " + element_to_click)
    driver.find_element_by_xpath(element_to_click).click()

def get_last_clicked_xpath(line_number):
    last_clicked_xpath = ""
    for i in range(line_number):
        if dictionary_of_action_selector_variable_per_row[i].get() == "Click (Xpath)":
            last_clicked_xpath = dictionary_of_action_input_per_row[i].get()
    return last_clicked_xpath

def enter_text(driver, line_number):
    last_clicked_xpath = get_last_clicked_xpath(line_number)
    input_text = dictionary_of_action_input_per_row[line_number].get()
    insert_text("Typing " + '"' + input_text + '"' + " in element " + '"' + last_clicked_xpath + '"')
    driver.find_element_by_xpath(last_clicked_xpath).send_keys(input_text)

def create_browser():
    options = webdriver.ChromeOptions()
    if display_browser_var.get() == 0:
        options.add_argument('headless')
        insert_text('Opening "hidden" browser.')
    selected_language = ["en-us", "en-gb", "de-de", "fr-fr", "es-es", "it-it", "pl-pl", "pt-br", "pt-pt", "ru-ru", "ko-kr", "zh-tw", "ja-jp", "th-th"]
    options.add_experimental_option('prefs', {'intl.accept_languages': selected_language[9]})
    pc_browser = "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36"
    pc_browser_width = 1920
    pc_browser_height = 1080
    mobile_browser = 'user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 10_3 like Mac OS X) AppleWebKit/602.1.50 (KHTML, like Gecko) CriOS/56.0.2924.75 Mobile/14E5239e Safari/602.1'
    mobile_browser_width = 768
    mobile_browser_height = 1024
    options.add_argument(pc_browser)
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(pc_browser_width, pc_browser_height)
    driver.implicitly_wait(10)
    return driver

def take_screenshot(driver, line_number, language):
    screenshot_name = dictionary_of_action_input_per_row[line_number].get() + "_" + language[1:] + ".png"
    insert_text("Taking screenshot: " + screenshot_name)
    whole_page = driver.find_element_by_tag_name('body')
    whole_page.screenshot(screenshot_name)

def single_action(driver, line_number, language):
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Open (URL)":
        try:
            go_to(driver, line_number, language)
        except:
            insert_text("Was unable to open this URL.")
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Click (Xpath)":
        try:
            click_element(driver, line_number)
        except:
            insert_text("Was unable to click the element.")
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Send (text)":
        try:
            enter_text(driver, line_number)
        except:
            insert_text("Was unable to send text.")
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Screenshot (save as)":
        try:
            take_screenshot(driver, line_number, language)
        except:
            insert_text("Was unable to save screenshot.")
    if dictionary_of_action_selector_variable_per_row[line_number].get() == "Wait (seconds)":
        try:
            time_to_wait = eval(dictionary_of_action_input_per_row[line_number].get())
            insert_text("Waiting " + str(time_to_wait) + " second(s).")
            time.sleep(time_to_wait)
        except:
            insert_text("Input not recognized. Waiting 1 second.")
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
        languages.append(" default url")
    def slow_magic():
        insert_text("Starting the process!")
        insert_text("Please, don't interact with interface until done. You can minimize this window.")
        for language in languages:
            driver = create_browser()
            for line_number in range(row_counter):
                insert_text("Performing step " + str(line_number + 1) + " for " + language[1:])
                single_action(driver, line_number, language)
            driver.quit()
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

def open_apps_folder():
    os.startfile(current_directory + "\\", 'open') 

def save_testing_sequence(entry_test_sequence_name, popup_provide_test_sequence_name):
    global row_counter
    wb = Workbook()
    ws = wb.active
    for line_number in range(row_counter):
        ws.cell(row=(line_number + 1), column=1, value=dictionary_of_action_selector_variable_per_row[line_number].get())
        ws.cell(row=(line_number + 1), column=2, value=dictionary_of_action_input_per_row[line_number].get())
    test_sequence_name = entry_test_sequence_name.get()
    if ".xlsx" not in test_sequence_name:
        test_sequence_name += ".xlsx"
    wb.save(test_sequence_name)
    popup_provide_test_sequence_name.destroy()

def get_main_window_of_gui_postion():
    x_coordinate = main_window_of_gui.winfo_x()
    y_coordinate = main_window_of_gui.winfo_y()
    return ("230x80" + "+" + (str(x_coordinate) + "+" + str(y_coordinate)))

def save_testing_sequence_popup():
    popup_provide_test_sequence_name = Toplevel()
    popup_provide_test_sequence_name.geometry(get_main_window_of_gui_postion())
    popup_provide_test_sequence_name.wm_attributes("-topmost", 1)
    popup_provide_test_sequence_name.wm_title("Test sequence name")
    label_test_sequence_name = Label(popup_provide_test_sequence_name, text = "Please, provide the name of this test run:")
    label_test_sequence_name.pack()
    entry_test_sequence_name = Entry(popup_provide_test_sequence_name, width=30)
    entry_test_sequence_name.pack()
    entry_test_sequence_name.focus()
    button_save_test_sequence_name = Button(popup_provide_test_sequence_name, text="Save test sequence", width = 30, height = 3, command = lambda: save_testing_sequence(entry_test_sequence_name, popup_provide_test_sequence_name))
    button_save_test_sequence_name.pack()
    popup_provide_test_sequence_name.mainloop()

def find_the_row_of_the_next_empty_cell(test_sequence_file_name):
    number_of_lines = 1
    wb = load_workbook(test_sequence_file_name)
    ws = wb.active
    cell_to_check = ws['A' + str(number_of_lines)]
    while cell_to_check.value != None:
        number_of_lines = number_of_lines + 1
        cell_to_check = ws['A' + str(number_of_lines)]
    return number_of_lines, ws

def build_test_sequence_with_import(test_sequence_file_name):
    global row_counter
    while row_counter > 0:
        remove_action_row()
    number_of_lines, ws = find_the_row_of_the_next_empty_cell(test_sequence_file_name)
    for line in range(number_of_lines - 1):
        add_action_row()
        dictionary_of_action_selector_variable_per_row[line].set(ws.cell(row=line+1, column=1).value)
        dictionary_of_action_input_per_row[line].insert(0, ws.cell(row=line+1, column=2).value)

def load_testing_sequence():
    file_path = askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    test_sequence_file_name = os.path.basename(file_path)
    build_test_sequence_with_import(test_sequence_file_name)
    


main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("Screen-shooter v31/01/2019")
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
language_all_toggle.grid(row = 0, column = 9)
language_en_gb_toggle = Checkbutton(main_window_of_gui, text="en-GB", variable=language_en_gb_var)
language_en_gb_toggle.grid(row = 0, column = 0, columnspan = 2)
language_de_de_toggle = Checkbutton(main_window_of_gui, text="de_DE", variable=language_de_de_var)
language_de_de_toggle.grid(row = 0, column = 2)
language_fr_fr_toggle = Checkbutton(main_window_of_gui, text="fr-FR", variable=language_fr_fr_var)
language_fr_fr_toggle.grid(row = 0, column = 3)
language_es_es_toggle = Checkbutton(main_window_of_gui, text="es-ES", variable=language_es_es_var)
language_es_es_toggle.grid(row = 0, column = 4)
language_it_it_toggle = Checkbutton(main_window_of_gui, text="it-IT", variable=language_it_it_var)
language_it_it_toggle.grid(row = 0, column = 5)
language_pt_br_toggle = Checkbutton(main_window_of_gui, text="pt-BR", variable=language_pt_br_var)
language_pt_br_toggle.grid(row = 0, column = 6)
language_pl_pl_toggle = Checkbutton(main_window_of_gui, text="pl-PL", variable=language_pl_pl_var)
language_pl_pl_toggle.grid(row = 0, column = 7)
language_ru_ru_toggle = Checkbutton(main_window_of_gui, text="ru-RU", variable=language_ru_ru_var)
language_ru_ru_toggle.grid(row = 0, column = 8)

display_browser_toggle = Checkbutton(main_window_of_gui, text="Display browser", variable=display_browser_var)
display_browser_toggle.grid(row = 0, column = 10)

button_add_action_row = Button(main_window_of_gui, text = "Add row", command = add_action_row)
button_add_action_row.grid(row = 1, column = 10)
button_remove_action_row = Button(main_window_of_gui, text = "Remove row", command = remove_action_row)
button_remove_action_row.grid(row = 1, column = 11)
button_perform_actions = Button(main_window_of_gui, text = "Perform actions", width = 16, height = 3, command = perform_actions)
button_perform_actions.grid(row = 2, column = 10, rowspan = 3, columnspan = 2)

text_box = Listbox(main_window_of_gui, height=8)
text_box.grid(column=0, row=100, columnspan=13, sticky=(N,W,E,S))  # columnspan − How many columns widgetoccupies; default 1.
main_window_of_gui.grid_columnconfigure(0, weight=1)
main_window_of_gui.grid_rowconfigure(13, weight=1)
#scroll bar
my_scrollbar = Scrollbar(main_window_of_gui, orient=VERTICAL, command=text_box.yview)
my_scrollbar.grid(column=13, row=100, sticky=(N,S))
#attaching scroll bar to text box
text_box['yscrollcommand'] = my_scrollbar.set

button_open_the_folder = Button(main_window_of_gui, text = "Open the folder", width = 16, height = 2, command = open_apps_folder)
button_open_the_folder.grid(row = 5, column = 10, rowspan = 2, columnspan = 2)

button_save_testing_sequence = Button(main_window_of_gui, text = "Save tests", width = 8, height = 2, command = save_testing_sequence_popup)
button_save_testing_sequence.grid(row = 7, column = 10, rowspan = 2, columnspan = 1)

button_load_testing_sequence = Button(main_window_of_gui, text = "Load tests", width = 8, height = 2, command = load_testing_sequence)
button_load_testing_sequence.grid(row = 7, column = 11, rowspan = 2, columnspan = 1)

main_window_of_gui.mainloop()

#driver.quit()
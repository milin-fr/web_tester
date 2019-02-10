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
from tkinter import Label, Button, Entry, Checkbutton, OptionMenu, Canvas, Frame
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

SCREENSHOT_FOLDER_NAME = ""

stop_test = 0

row_counter = 0

test_sequence_source_excel = ""

dictionary_of_action_index_per_row = {}
dictionary_of_action_selector_per_row = {}
dictionary_of_action_selector_variable_per_row = {}
dictionary_of_action_input_per_row = {}


list_of_actions = ["Open (URL)", "Click (Xpath)", "Send (text)", "Screenshot (save as)", "Wait (seconds)", "Open (URL) +", "Click (Xpath) +", "Send (text) +", "Screenshot (save as) +"]
list_of_browsers = ["PC browser 1920x1080", "Mobile browser 768x1204"]


def get_files_in_script_directory():
    '''Get file names in directory'''
    file_names = []
    for root, dirs, files in os.walk(current_directory):
        for filename in files:
            file_names.append(filename)
    return file_names

def create_folder_for_screenshots(dictionary_of_sequences):
    global SCREENSHOT_FOLDER_NAME
    time_stamp_minutes = str(time.strftime("%m-%d %Hh%M"))
    need_to_create_folder = 0
    for sequence in dictionary_of_sequences:
        for action, value in dictionary_of_sequences[sequence].items():
            if dictionary_of_sequences[sequence][action]["action"] == "Screenshot (save as)" or dictionary_of_sequences[sequence][action]["action"] == "Screenshot (save as) +":
                need_to_create_folder = 1
    if need_to_create_folder == 1:
        try:
            SCREENSHOT_FOLDER_NAME = "./Screenshots " + time_stamp_minutes + "/"
            os.makedirs(SCREENSHOT_FOLDER_NAME)
        except:
            time_stamp_seconds = str(time.strftime("%Hh%Ms%S"))
            SCREENSHOT_FOLDER_NAME = "./Screenshots " + time_stamp_seconds + "/"
            os.makedirs(SCREENSHOT_FOLDER_NAME)
        for sequence_index in dictionary_of_sequences:
            os.makedirs(SCREENSHOT_FOLDER_NAME + "/" + str(sequence_index) + "/")

def look_up_action_value_from_excel(single_value, test_reiteration):
    wb = load_workbook(test_sequence_source_excel)
    ws = wb.worksheets[1]
    needed_column = 0
    needed_row = test_reiteration + 2
    number_of_columns = 1
    columns_to_check = ws.cell(row=1, column=number_of_columns)
    while columns_to_check.value != None:
        if columns_to_check.value == single_value:
            needed_column = number_of_columns
        number_of_columns = number_of_columns + 1
        columns_to_check = ws.cell(row=1, column=number_of_columns)
    needed_value_cell = ws.cell(row=needed_row, column=needed_column)
    while needed_value_cell.value == None:
        needed_row -= 1
        needed_value_cell = ws.cell(row=needed_row, column=needed_column)
    return needed_value_cell.value


def build_single_test_sequence(test_reiteration):
    list_of_complex_inputs = ["Open (URL) +", "Click (Xpath) +", "Send (text) +", "Screenshot (save as) +", "Language +", "Browser +"]
    dictionary_of_single_test_sequence = {}
    for i in range(row_counter):
        dictionary_of_single_test_sequence[i] = {}
        single_action = dictionary_of_action_selector_variable_per_row[i].get()
        dictionary_of_single_test_sequence[i]["action"] = single_action
        single_value = dictionary_of_action_input_per_row[i].get()
        if single_action in list_of_complex_inputs:
            complex_value = look_up_action_value_from_excel(single_value, test_reiteration)
            dictionary_of_single_test_sequence[i]["value"] = complex_value
        else:
            dictionary_of_single_test_sequence[i]["value"] = single_value
    return dictionary_of_single_test_sequence

def check_if_excel_file_is_selected_and_create_one_if_not():
    if test_sequence_source_excel == "":
        save_testing_sequence_popup()

def get_the_number_of_test_sequences():
    number_of_test_sequences = 0
    selected_language = language_option_var.get()
    if selected_language == "language +":
        check_if_excel_file_is_selected_and_create_one_if_not()
    try:
        wb = load_workbook(test_sequence_source_excel)
        ws = wb.worksheets[1]
        number_of_columns = 1
        columns_to_check = ws.cell(row=1, column=number_of_columns)
        while columns_to_check.value != None:
            number_of_columns = number_of_columns + 1
            columns_to_check = ws.cell(row=1, column=number_of_columns)
        for column_to_check in range(1, number_of_columns):
            number_of_rows = 1
            rows_to_check = ws.cell(row=number_of_rows, column=column_to_check)
            while rows_to_check.value != None:
                number_of_rows = number_of_rows + 1
                rows_to_check = ws.cell(row=number_of_rows, column=column_to_check)
            if number_of_test_sequences < number_of_rows:
                number_of_test_sequences = number_of_rows
        number_of_test_sequences -= 1
    except:
        number_of_test_sequences = 1
    return number_of_test_sequences



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
    action_row = (row_counter) % 20
    action_column = row_counter // 20  * 3 
    dictionary_of_action_index_per_row[row_counter] = Label(main_window_of_gui, text = str(row_counter + 1))
    dictionary_of_action_index_per_row[row_counter].grid(row = action_row + 3, column = action_column)
    dictionary_of_action_selector_variable_per_row[row_counter] = StringVar()
    dictionary_of_action_selector_variable_per_row[row_counter].set(list_of_actions[0])
    dictionary_of_action_selector_per_row[row_counter] = OptionMenu(main_window_of_gui, dictionary_of_action_selector_variable_per_row[row_counter], *list_of_actions)
    dictionary_of_action_selector_per_row[row_counter].config(width=18)
    dictionary_of_action_selector_per_row[row_counter].grid(row = action_row + 3, column = action_column + 1, columnspan = 1)
    dictionary_of_action_input_per_row[row_counter] = Entry(main_window_of_gui)
    dictionary_of_action_input_per_row[row_counter].config(width=55)
    dictionary_of_action_input_per_row[row_counter].grid(row = action_row + 3, column = action_column + 2, columnspan = 1)
    row_counter += 1


def go_to(driver, value):
    insert_text("Opening " + value)
    driver.get(value)

def click_element(driver, value):
    insert_text("Clicking " + value)
    driver.find_element_by_xpath(value).click()

def get_last_clicked_xpath(dictionary_of_current_test_itiration, action_index):
    last_clicked_xpath = ""
    for i in range(action_index):
        if dictionary_of_current_test_itiration[i]["action"] == "Click (Xpath)" or dictionary_of_current_test_itiration[i]["action"] == "Click (Xpath) +":
            last_clicked_xpath = dictionary_of_current_test_itiration[i]["value"]
    return last_clicked_xpath

def enter_text(driver, value, dictionary_of_current_test_itiration, action_index):
    last_clicked_xpath = get_last_clicked_xpath(dictionary_of_current_test_itiration, action_index)
    input_text = value
    insert_text("Typing " + '"' + input_text + '"' + " in element " + last_clicked_xpath)
    driver.find_element_by_xpath(last_clicked_xpath).send_keys(input_text)

def create_browser(test_reiteration):
    pc_browser = "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36"
    pc_browser_width = 1920
    pc_browser_height = 1080
    mobile_browser = 'user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 10_3 like Mac OS X) AppleWebKit/602.1.50 (KHTML, like Gecko) CriOS/56.0.2924.75 Mobile/14E5239e Safari/602.1'
    mobile_browser_width = 768
    mobile_browser_height = 1024
    options = webdriver.ChromeOptions()
    if display_browser_var.get() == 0:
        options.add_argument('headless')
        insert_text('Opening "hidden" browser.')
    selected_language = language_option_var.get()
    if selected_language == "language +":
        selected_language = look_up_action_value_from_excel("language", test_reiteration)
        print(selected_language)
    options.add_experimental_option('prefs', {'intl.accept_languages': selected_language})
    insert_text('Setting browser default language to ' + selected_language + ".")
    if browser_type_var.get() == "PC browser 1920x1080":
        selected_browser = pc_browser
        selected_width = pc_browser_width
        selected_height = pc_browser_height
    if browser_type_var.get() == "Mobile browser 768x1204":
        selected_browser = mobile_browser
        selected_width = mobile_browser_width
        selected_height = mobile_browser_height
    options.add_argument(selected_browser)
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(selected_width, selected_height)
    driver.implicitly_wait(10)
    return driver

def take_screenshot(driver, value, action_index, test_reiteration):
    global SCREENSHOT_FOLDER_NAME
    time_stamp_seconds = str(time.strftime("%Hh%Ms%S"))
    screenshot_name = value
    if ".png" not in screenshot_name:
        screenshot_name += ".png"
    screenshot_name = time_stamp_seconds + " " + str(action_index) + " " + screenshot_name
    insert_text("Taking screenshot: " + screenshot_name)
    screenshot_name = SCREENSHOT_FOLDER_NAME + screenshot_name
    whole_page = driver.find_element_by_tag_name('body')
    whole_page.screenshot(screenshot_name)

def single_action(driver, action_value_tuple, dictionary_of_current_test_itiration, action_index, test_reiteration):
    action = action_value_tuple["action"]
    value = action_value_tuple["value"]
    if action == "Open (URL)" or action == "Open (URL) +":
        try:
            go_to(driver, value)
        except:
            insert_text("Was unable to open " + value)
    if action == "Click (Xpath)" or action == "Click (Xpath) +":
        try:
            click_element(driver, value)
        except:
            insert_text("Was unable to click " + value)
    if action == "Send (text)" or action == "Send (text) +":
        try:
            enter_text(driver, value, dictionary_of_current_test_itiration, action_index)
        except:
            insert_text("Was unable to write " + value)
    if action == "Screenshot (save as)" or action == "Screenshot (save as) +":
        try:
            take_screenshot(driver, value, action_index, test_reiteration)
        except:
            insert_text("Was unable to take screenshot " + value)
    if action == "Wait (seconds)":
        time.sleep(int(value))

"""
    def slow_magic():
    executing = Thread(target=slow_magic)
    executing.start()
"""

def perform_actions(dictionary_of_current_test_itiration, test_reiteration):
    driver = create_browser(test_reiteration)
    for action_index, action_value_tuple in dictionary_of_current_test_itiration.items():
        single_action(driver, action_value_tuple, dictionary_of_current_test_itiration, action_index, test_reiteration)
    driver.quit()


def insert_text(text):
    text_box.insert('end', text)
    text_box.see("end")

def open_apps_folder():
    os.startfile(current_directory + "\\", 'open') 

def save_testing_sequence(entry_test_sequence_name, popup_provide_test_sequence_name):
    global test_sequence_source_excel
    global row_counter
    test_sequence_name = entry_test_sequence_name.get()
    wb = Workbook()
    ws = wb.worksheets[0]
    for line_number in range(row_counter):
        ws.cell(row=(line_number + 1), column=1, value=dictionary_of_action_selector_variable_per_row[line_number].get())
        ws.cell(row=(line_number + 1), column=2, value=dictionary_of_action_input_per_row[line_number].get())
    list_of_single_languages = ["en-us", "en-gb", "de-de", "fr-fr", "es-es", "it-it", "pl-pl", "pt-br", "pt-pt", "ru-ru", "ko-kr", "zh-tw", "ja-jp", "th-th"]
    selected_language = language_option_var.get()
    if selected_language == "language +":
        wb.create_sheet()
        ws = wb.worksheets[1]
        ws.cell(row=1, column=1, value="language")
        next_row = 2
        for single_language in list_of_single_languages:
            ws.cell(row=next_row, column=1, value=single_language)
            next_row +=1
    if ".xlsx" not in test_sequence_name:
        test_sequence_name += ".xlsx"
    wb.save(test_sequence_name)
    test_sequence_source_excel = test_sequence_name
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

def find_the_row_of_the_next_empty_cell():
    number_of_lines = 1
    wb = load_workbook(test_sequence_source_excel)
    ws = wb.worksheets[0]
    cell_to_check = ws.cell(row = number_of_lines, column = 1)
    while cell_to_check.value != None:
        number_of_lines = number_of_lines + 1
        cell_to_check = ws.cell(row = number_of_lines, column = 1)
    return number_of_lines, ws

def build_test_sequence_with_import():
    global row_counter
    while row_counter > 0:
        remove_action_row()
    number_of_lines, ws = find_the_row_of_the_next_empty_cell()
    for line in range(number_of_lines - 1):
        add_action_row()
        dictionary_of_action_selector_variable_per_row[line].set(ws.cell(row=line+1, column=1).value)
        dictionary_of_action_input_per_row[line].insert(0, ws.cell(row=line+1, column=2).value)

def choose_test_sequence_source_excel():
    global test_sequence_source_excel
    test_sequence_source_excel = askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    build_test_sequence_with_import()

def stop_the_test():
    global stop_test
    stop_test = 1

#def running_threads(ai_response):
#threads = []
#for i in range(10):
#t = Thread(target = get_article_name, args = (ai_response, i))
#threads.append(t)
#t.start()
#for t in threads:t.join()
#
#def get_article_name(ai_response, i):
#try:
#page_content = request.urlopen("https://battle.net/support/article/" + str(ai_response["suggestedArticles"][i]["articleNumber"]))
#soup = BeautifulSoup(page_content, 'html.parser')
#article_name[order]= soup.find(class_="section-header").text
#except:
#article_name[order]= "Internal article " + str(ai_response["suggestedArticles"][i]["articleNumber"])


def run_selected_actions():
    def slow_magic():
        global stop_test
        list_of_sequences = []
        stop_test = 0
        text_box.delete(0,'end')
        insert_text("Starting the process!")
        insert_text("Please, don't interact with interface until done. You can minimize this window.")
        dictionary_of_sequences = {}
        number_of_test_sequences = get_the_number_of_test_sequences()
        for test_reiteration in range(number_of_test_sequences):
            dictionary_of_sequences[test_reiteration] = build_single_test_sequence(test_reiteration)
        create_folder_for_screenshots(dictionary_of_sequences)
        for test_reiteration in range(number_of_test_sequences):
            if stop_test == 0:
                t = Thread(target = perform_actions, args = (dictionary_of_sequences[test_reiteration], test_reiteration))
                list_of_sequences.append(t)
                t.start()
                #perform_actions(dictionary_of_sequences[test_reiteration], test_reiteration)
            else:
                insert_text("/!\\ Sequence aborted /!\\")
        for t in list_of_sequences:t.join()
        insert_text("****************** ALL ACTIONS PERFORMED ******************")
    executing = Thread(target=slow_magic)
    executing.start()

"""
    def slow_magic():
    executing = Thread(target=slow_magic)
    executing.start()
"""

main_window_of_gui = tkinter.Tk()
main_window_of_gui.title("Screen-shooter v05/02/2019")
main_window_of_gui.wm_attributes("-topmost", 1)

language_option_var = StringVar()
language_option_list = ["en-us", "en-gb", "de-de", "fr-fr", "es-es", "it-it", "pl-pl", "pt-br", "pt-pt", "ru-ru", "ko-kr", "zh-tw", "ja-jp", "th-th", "language +"]

browser_type_var = StringVar()
display_browser_var = IntVar()

display_browser_toggle = Checkbutton(main_window_of_gui, text="Display browser", variable=display_browser_var)
display_browser_toggle.grid(row = 0, column = 0, columnspan = 2)

label_language = Label(main_window_of_gui, text = "Select default browser language: ")
label_language.grid(row = 1, column = 0, columnspan = 2)

language_option_var.set(language_option_list[0])
drop_down_language = OptionMenu(main_window_of_gui, language_option_var, *language_option_list)
drop_down_language.config(width = 22)
drop_down_language.grid(row = 1, column = 2, columnspan = 1)

label_browser = Label(main_window_of_gui, text = "Select browser type: ")
label_browser.grid(row = 2, column = 0, columnspan = 2)

browser_type_var.set(list_of_browsers[0])
drop_down_browser_type = OptionMenu(main_window_of_gui, browser_type_var, *list_of_browsers)
drop_down_browser_type.config(width = 22)
drop_down_browser_type.grid(row = 2, column = 2, columnspan = 1)

button_add_action_row = Button(main_window_of_gui, text = "Add row", width = 16, height = 3, command = add_action_row)
button_add_action_row.grid(row = 0, column = 2, columnspan = 2)
button_remove_action_row = Button(main_window_of_gui, text = "Remove row", width = 16, height = 3, command = remove_action_row)
button_remove_action_row.grid(row = 0, column = 4, columnspan = 1)
button_perform_actions = Button(main_window_of_gui, text = "Perform actions", width = 16, height = 3, command = run_selected_actions) #perform_actions
button_perform_actions.grid(row = 0, column = 5, rowspan = 1, columnspan = 1)

text_box = Listbox(main_window_of_gui, height=8)
text_box.grid(column=0, row=100, columnspan=13, sticky=(N,W,E,S))  # columnspan − How many columns widgetoccupies; default 1.
main_window_of_gui.grid_columnconfigure(0, weight=1)
main_window_of_gui.grid_rowconfigure(13, weight=1)
#scroll bar
my_scrollbar = Scrollbar(main_window_of_gui, orient=VERTICAL, command=text_box.yview)
my_scrollbar.grid(column=13, row=100, sticky=(N,S))
#attaching scroll bar to text box
text_box['yscrollcommand'] = my_scrollbar.set

button_open_the_folder = Button(main_window_of_gui, text = "Open the folder", width = 16, height = 3, command = open_apps_folder)
button_open_the_folder.grid(row = 0, column = 6, rowspan = 1, columnspan = 1)

button_save_testing_sequence = Button(main_window_of_gui, text = "Save tests", width = 16, height = 3, command = save_testing_sequence_popup)
button_save_testing_sequence.grid(row = 0, column = 7, rowspan = 1, columnspan = 1)

button_load_testing_sequence = Button(main_window_of_gui, text = "Load tests", width = 16, height = 3, command = choose_test_sequence_source_excel)
button_load_testing_sequence.grid(row = 0, column = 8, rowspan = 1, columnspan = 1)

button_stop_the_test = Button(main_window_of_gui, text = "Cancel", width = 16, height = 3, command = stop_the_test)
button_stop_the_test.grid(row = 0, column = 9, rowspan = 1, columnspan = 1)

main_window_of_gui.mainloop()
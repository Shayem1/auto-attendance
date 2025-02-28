from tkinter import *
import tkinter as tk
import customtkinter
import xlsxwriter
import time
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By

# Load the Excel workbook
file_path = "IBMA.xlsx"
wb = openpyxl.load_workbook(file_path)

# Select the active sheet
sheet = wb.active

# Create an empty list to store the data from the first column
students = []

# Loop through the rows and extract data from the first column (column A)
for row in sheet.iter_rows(min_row=1, max_col=1, values_only=True):  # Skipping the header row
    students.append(row[0]) 

wb.close()


#button timeouts
def active():
    global cont_button
    cont_button.configure(state=NORMAL)


def phase_3():
    global driver, cont_button, label, students

    label.configure(text="Go to Page 4 and press continue and add to\nmerge documents, once completed, repeat\nfor all pages", font=("arial",20))

    data_list = []
    checkbox = []
    index_list = []

    found_data = driver.find_elements(By.XPATH ,'//*[@id="myForm"]/table/tbody/tr/td[4]') 
    checkbox_data = driver.find_elements(By.XPATH ,'/html/body/div[1]/div[4]/div/form[2]/table/tbody/tr/td[1]')

    for i in found_data:
        data_list.append(i.text)

    for index, value in enumerate(data_list):
        if value in students:
            index_list.append(index)

    for i in checkbox_data:
            checkbox.append(i)

    for i in index_list:
        checkbox[i].click()


def phase_2():

    #move to
    global status, driver, label, cont_button
    status = 2

    #opens the next webpage
    driver.get("https://ibma.app.axcelerate.com/management/management2/ContactsSearch.cfm")

    #next instructions
    label.configure(text="Please select optional id\nand sort by optional id", font=("arial",20))
    

    cont_button.configure(text="Continue", font=("arial",30), corner_radius=30, state=DISABLED,  command= lambda: phase_3())
    

    #ensures user does the action
    GUI.after(5000, active)

def phase_1():
    #move to first stage
    global status, label, cont_button, driver
    status = 1

    #opens chrome page
    driver = webdriver.Chrome()
    driver.get("https://ibma.app.axcelerate.com/management/")
    

    #adds instructions for the next stage
    label.configure(text="Please Login to any account\n with access to merge mailout groups\n\nPress continue AFTER you login")
    

    cont_button.configure(state=DISABLED,  command= lambda: phase_2())
    

    GUI.after(20000, active)

#helps code remember what phase it is on
status = 0

#Gui setup
GUI = customtkinter.CTk()
GUI.title("Tasker")
GUI.geometry("450x200")
GUI.minsize(width=450, height=200)
GUI.maxsize(width=450, height=200)
customtkinter.set_appearance_mode("dark")

#setting up info text for phase 0
label = customtkinter.CTkLabel(master=GUI, text="Hello, Tasker will guide you on how\nto automatically make merge out groups\nin seconds from the attendance sheet", font=("arial",20))
label.pack(pady=15)

cont_button = customtkinter.CTkButton(master = GUI, text="Continue", font=("arial",30), corner_radius=30, command= lambda: phase_1())
cont_button.pack(pady=30)

GUI.mainloop()
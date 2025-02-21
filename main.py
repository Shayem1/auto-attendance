from tkinter import *
import tkinter as tk
import customtkinter
import xlsxwriter
import time
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver

#button timeouts
def active():
    global cont_button
    cont_button.configure(state=NORMAL)


def phase_3():
    global driver, cont_button, label
    pass


def phase_2():

    #move to
    global status, driver, label, cont_button
    status = 2

    #opens the next webpage
    driver.get("https://ibma.app.axcelerate.com/management/management2/ContactsSearch.cfm")

    #removes old info
    label.destroy()
    cont_button.destroy()

    #next instructions
    label.configure(master=GUI, text="Please select optional id\nand sort by optional id", font=("arial",20))
    label.pack(pady=15)

    cont_button.configure(text="Continue", font=("arial",30), corner_radius=30, state=DISABLED,  command= lambda: phase_3())
    cont_button.pack(pady=15)

    #ensures user does the action
    GUI.after(5000, active)

def phase_1():
    #move to first stage
    global status, label, cont_button, driver
    status = 1

    #opens chrome page
    driver = webdriver.Chrome()
    driver.get("https://ibma.app.axcelerate.com/management/")
    
    #removes objects from previous stage
    cont_button.destroy()
    label.destroy()

    #adds instructions for the next stage
    label.configure(text="Please Login to any account\n with access to merge mailout groups\n\nPress continue AFTER you login", font=("arial",20))
    label.pack(pady=15)

    cont_button.configure(text="Continue", font=("arial",30), corner_radius=30, state=DISABLED,  command= lambda: phase_2())
    cont_button.pack(pady=15)

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

global label, cont_button

#setting up info text for phase 0
label = customtkinter.CTkLabel(master=GUI, text="Hello, Tasker will guide you on how\nto automatically make merge out groups\nin seconds from the attendance sheet", font=("arial",20))
label.pack(pady=15)

cont_button = customtkinter.CTkButton(master = GUI, text="Continue", font=("arial",30), corner_radius=30, command= lambda: phase_1())
cont_button.pack(pady=30)

GUI.mainloop()
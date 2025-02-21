from tkinter import *
import tkinter as tk
import customtkinter
import xlsxwriter
import time
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver


def phase_2():

    #move to 
    status = 2



def phase_1():
    
    #move to first stage
    global status, label2, cont_button1
    status = 1

    #opens chrome page
    driver = webdriver.Chrome()
    driver.get("https://ibma.app.axcelerate.com/management/")
    
    #removes objects from previous stage
    cont_button.destroy()
    label1.destroy()

    #adds instructions for the next stage
    label2 = customtkinter.CTkLabel(master=GUI, text="Please Login to any account\n with access to merge mailout groups\n\nPress continue AFTER you login", font=("arial",20))
    label2.pack(pady=15)

    cont_button1 = customtkinter.CTkButton(master = GUI, text="Continue", font=("arial",30), corner_radius=30, command= lambda: phase_2())
    cont_button1.pack(pady=15)

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
label1 = customtkinter.CTkLabel(master=GUI, text="Hello, Tasker will guide you on how\nto automatically make merge out groups\nin seconds from the attendance sheet", font=("arial",20))
label1.pack(pady=15)

cont_button = customtkinter.CTkButton(master = GUI, text="Continue", font=("arial",30), corner_radius=30, command= lambda: phase_1())
cont_button.pack(pady=30)

GUI.mainloop()
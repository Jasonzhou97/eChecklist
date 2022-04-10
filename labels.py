from tkinter import *

def spacing(x,y):
    Label(text="     ").grid(row=x,column=y)

def create(text,x,y):
    Label(text=text).grid(row=x,column=y)



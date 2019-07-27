from functools import partial
from openpyxl import Workbook
from os.path import *
from pathlib import *
from PIL import Image
from six.moves import input
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import csv
import ftransc
import numpy as np
import os
import PIL
import time

def clicked():
    Error.set("")
    oldFiletype = str(fromFiletype.get())
    newFiletype = str(toFiletype.get())
    getEnding(newFiletype)

    if (oldFiletype == newFiletype):
        Error.set("You can't convert a " +oldFiletype +" to a " +newFiletype +"!")
        ErrorMessage()
    elif (oldFiletype == ".tai") and (newFiletype == ".csv"):
        taiToCsv()
    elif (oldFiletype == ".csv") and (newFiletype == ".xlsx"):
        csvToXlsx()
    elif (oldFiletype == ".csv") and (newFiletype == ".txt"):
        csvToTxt()
    elif (oldFiletype == ".tai") and (newFiletype == ".xlsx"):
        taiToXlsx()
    elif (oldFiletype == ".tai") and (newFiletype == ".txt"):
        taiToTxt()
    elif (oldFiletype == ".png") and (newFiletype == ".jpg"):
        pngToJpg()
    elif (oldFiletype == ".png") and (newFiletype == ".tiff"):
        pngToTiff()
    elif (oldFiletype == ".tiff") and (newFiletype == ".png"):
        tiffToPng()
    else:
        Error.set("It's not possible to convert a " +oldFiletype + " to a " +newFiletype +"!")
        ErrorMessage()
        
def csvToXlsx():
    entry_file = str(pathway.get()+filename.get())
    product_file = str(filenameEntered.get())

    checking_entry_file = entry_file.endswith(".csv")#Checking the ending of the filenames to prevent an error
    checking_product_file = product_file.endswith(".xlsx")

    if checking_entry_file == False:
        entry_file = entry_file +".csv"
    if checking_product_file == False:
        product_file = product_file +".xlsx"

    try:
        with open(entry_file) as f:
            for row in csv.reader(f):
                ws.append(row)
        wb.save(product_file)
        loadingAnimation()
        reset()
    except Exception as e:
        E = str(e)
        if ("[Errno 2] No such file or directory" in E):
            Error.set("No such file or direcotry!")
            ErrorMessage()
        else:
            Error.set("Unknown error occured!")
            ErrorMessage()

def destroyToplevel(window): #destroys the new windows(Help, ErrorMessage, Settings, loadingAnimation)
                          window.destroy()

def ErrorMessage(): #Creating a new window to show the error message
    error_window = Toplevel()
    error_window.config(background = backgroundColor)
    error_window.title("An error occured!")
    error_label = Label(error_window, background = backgroundColor, foreground = "red", textvariable = Error)
    error_label.grid(column = 0, row = 0)
    error_button = Button(error_window, background = backgroundColor, foreground = foregroundColor, text = "Okay", command = partial (destroyToplevel, error_window))
    error_button.grid(column = 0, row = 1)
    error_window.lift()

def getEnding(file):
    if file.endswith(".csv"):
        fromFiletypeChoose.current(0)
    elif file.endswith(".tai"):
        fromFiletypeChoose.current(3)
    elif file.endswith(".txt"):
        fromFiletypeChoose.current(5)
    elif file.endswith(".xlsx"):
        fromFiletypeChoose.current(6)
    elif file.endswith(".jpg"):
        fromFiletypeChoose.current(1)
    elif file.endswith(".png"):
        fromFiletypeChoose.current(2)
    elif file.endswith(".tiff"):
        fromFiletypeChoose.current(4)
    else:
        fromFiletypeChoose.current(0)

def Help(): #creating a new window to show the help message
    os.startfile("User Support.html")

def loadingAnimation():
    divisions = 100
    currentValue = 0
    def progress(currentValue):
            progressbar["value"] = currentValue
    animationwindow = Toplevel()
    animationwindow.config(background = backgroundColor)
    converted_label = Label(animationwindow, background = backgroundColor, foreground = "green", textvariable = Converted)
    converted_label.grid(column = 0, row = 1)
    okay_button = Button(animationwindow, background = backgroundColor, foreground = foregroundColor, text = "Okay", command = partial(destroyToplevel, animationwindow))
    okay_button.grid(column = 0, row = 2)
    progressbar = ttk.Progressbar(animationwindow, orient = "horizontal", length = 100, mode = "determinate")
    progressbar.grid(column = 0, row = 0)
    progressbar["value"] = currentValue
    progressbar["maximum"] = maxValue

    for i in range(divisions):
        currentValue = currentValue + 1
        progressbar.after(5, progress(currentValue))
        progressbar.update()
    Converted.set("Converted!")
    progress(0)

def pngToJpg():
    pathway_used = str(pathway.get())
    filename_entered = str(filename.get())
    filename_new = str(filenameEntered.get())

    checking_filename_new = filename_new.endswith(".png")#Checking the ending of the filenames to prevent an error
    checking_filename_entered = filename_entered.endswith(".jpg")

    if checking_filename_new == False:
        filename_new = filename_new +".png"
    if checking_filename_entered == False:
        filename_entered = filename_entered +".jpg"

    input_image = str(pathway_used + filename_entered)
    output_image = str(pathway_used + filename_new)

    jpg_file = PIL.Image.open(input_image)
    jpg_file.save(output_image)
    loadingAnimation()
    reset()

def pngToTiff():
    pathway_used = str(pathway.get())
    filename_entered = str(filename.get())
    filename_new = str(filenameEntered.get())

    checking_filename_new = filename_new.endswith(".tiff")#Checking the ending of the filenames to prevent an error
    checking_filename_entered = filename_entered.endswith(".png")

    if checking_filename_new == False:
        filename_new = filename_new +".tiff"
    if checking_filename_entered == False:
        filename_entered = filename_entered +".png"

    input_image = str(pathway_used + filename_entered)
    output_image = str(pathway_used + filename_new)

    png_file = PIL.Image.open(input_image)
    png_file.save(output_image)
    loadingAnimation()
    reset()

def reset():
    file = open("defaultResolution.txt","r")
    resolution = file.readline()
    file.close
    pathway.set(standart)
    filename.set(standart)
    filenameEntered.set(standart)
    Error.set(standart)
    Converted.set(standart)
    toFiletypeChoose.current(0)
    fromFiletypeChoose.current(0)
    window.geometry(resolution)


def selectFile():
    file_path = filedialog.askopenfilename()
    splitPath(file_path)

def setGeometry():
    geo = str(windowSize.get())
    if not("x" in geo):
        Error.set("Please enter a height!")
        ErrorMessage
    else:
        file = open("defaultResolution.txt","w")
        file.write(str(geo))
        file.close()
        window.geometry(geo)

def Settings():
    file = open("defaultResolution.txt","r")
    res = str(file.readline())
    file.close()
    windowSize.set(res)
    
    setting_window = Toplevel()
    setting_window.config(background = backgroundColor)
    label_1 = Label(setting_window, background = backgroundColor, foreground = foregroundColor, text = "Set your default geometry ('width'x'height'): ")
    label_1.grid(column = 0, row = 0)
    
    entry_geometry = Entry(setting_window, background = backgroundColor, foreground = foregroundColor, textvariable = windowSize)
    entry_geometry.grid(column = 1, row = 0)
    res = windowSize.get()
    
    button_saveGeometry = Button(setting_window, background = backgroundColor, foreground = foregroundColor, text = "Save", command = setGeometry)
    button_saveGeometry.grid(column = 2, row = 0)
    button_okay = Button(setting_window, background = backgroundColor, foreground = foregroundColor, text = "Okay", command = partial(destroyToplevel,setting_window))
    button_okay.grid(column = 0, row = 3)

def splitPath(path):
    lenghtPath = len(path)-1
    while path[lenghtPath] != "/":
        lenghtPath -= 1
    lenghtPath = lenghtPath +1
    div = len(path)-lenghtPath+1
    filenameNew = ""
    #for i in range(div-1):
    filenameNew = str(path[lenghtPath:])
        #filenameNew = filenameNew + path[lenghtPath]
        #lenghtPath += 1
    filepath = path.replace(filenameNew,"")
    pathway.set(filepath)
    filename.set(filenameNew)
    getEnding(filenameNew)

def taiToCsv():
    pathway_used = str(pathway.get())
    filename_entered = str(filename.get())
    filename_new = str(filenameEntered.get())

    checking_filename_new = filename_new.endswith(".csv")#Checking the ending of the filenames to prevent an error
    checking_filename_entered = filename_entered.endswith(".tai")

    if checking_filename_new == False:
        filename_new = filename_new +".csv"
    if checking_filename_entered == False:
        filename_entered = filename_entered +".tai"
    
    try:
     # Read the points from the original file, and write them to a new csv file
        with open(os.path.join(pathway_used, filename_entered), 'r') as csvfile:
            with open(join(pathway_used, filename_new),'w') as outcsv:
                spamwriter = csv.writer(outcsv, delimiter=',')
                lines_after = csvfile.readlines()[16:]
                spamreader = csv.reader(lines_after, delimiter='\t', quotechar='|')
                for line in spamreader:
                    if line[2]=="bad":
                        height = 99
                    else:
                        height = float(line[2])
                        grid_x = float(line[0])
                        grid_y = float(line[1])
                        spamwriter.writerow([grid_x, grid_y, height])
        loadingAnimation()
        reset()
    except Exception as e:
        E = str(e)
        print(E)
        if ("[Errno 2] No such file or directory" in E):
            Error.set("No such file or direcotry!")
            ErrorMessage()
        else:
            Error.set("Unknown error occured!")
            ErrorMessage()

def taiToTxt():
    taiToCsv()
    filename.set(filenameEntered.get())
    filenameNew = str(filenameEntered.get())
    filenameOld = filenameNew
    filenameNew.replace(".csv",".txt")
    filenameEntered.set(filenameNew)
    csvToTxt()
    path = str(pathway.get())
    file = path + filenameOld +".csv"
    os.remove(file)
    reset()

def taiToXlsx():
    taiToCsv()
    filename.set(filenameEntered.get())
    filenameNew = str(filenameEntered.get())
    filenameOld = filenameNew
    filenameNew.replace(".csv",".xlsx")
    filenameEntered.set(filenameNew)
    csvToXlsx()
    path = str(pathway.get())
    file = path + filenameOld +".csv"
    os.remove(file)
    reset()

def tiffToPng():
    pathway_used = str(pathway.get())
    filename_entered = str(filename.get())
    filename_new = str(filenameEntered.get())

    checking_filename_new = filename_new.endswith(".tiff")#Checking the ending of the filenames to prevent an error
    checking_filename_entered = filename_entered.endswith(".png")

    if checking_filename_new == False:
        filename_new = filename_new +".tiff"
    if checking_filename_entered == False:
        filename_entered = filename_entered +".png"

    input_image = str(pathway_used + filename_entered)
    output_image = str(pathway_used + filename_new)

    tiff_file = PIL.Image.open(input_image)
    tiff_file.save(output_image)
    loadingAnimation()
    reset()

def Quit():
    window.destroy()


try:
    file = open("defaultResolution.txt","r")#Getting the window geometry
    resolution = file.readline()
    file.close()
    foregroundColor = "black"
    backgroundColor = "white"
except:
    file = open("defaultResolution.txt","w+")
    resolution = "450x350"
    file.write(resolution)
    file.close()
    foregroundColor = "black"
    backgroundColor = "white"


wb = Workbook()
ws = wb.active
window = Tk()
window.title("Data Converter")
window.geometry(resolution)
window.config(background = backgroundColor)

window.grid_columnconfigure(0,weight=1)
window.grid_columnconfigure(1,weight=1)
window.grid_columnconfigure(2,weight=1)
window.grid_rowconfigure(0,weight=1)
window.grid_rowconfigure(1,weight=1)
window.grid_rowconfigure(2,weight=1)
window.grid_rowconfigure(3,weight=1)
window.grid_rowconfigure(4,weight=1)
window.grid_rowconfigure(5,weight=1)
window.grid_rowconfigure(6,weight=1)
window.grid_rowconfigure(7,weight=1)
window.grid_rowconfigure(8,weight=1)
window.resizable(False,False)

Converted = StringVar()
Error = StringVar()
filename = StringVar()
filenameMusic = StringVar()
filenameEntered = StringVar()
fromFiletype = StringVar()
help_message = StringVar()
pathway = StringVar()
toFiletype = StringVar()
windowSize = StringVar()
theme = IntVar()

standart = "" #Creating a 'standart' string to make the reset clearer
standartFiletypes = (".csv",".jpg",".png",".tai",".tiff",".txt",".xlsx") #Creating a 'filetypes' list
filetypesMusic = (".mp3",".mp4") #Creating a filetypes list for the music tab
currentValue = 0 #Setting the current value of the progressbar
maxValue = 100 #Setting the maximum value of the progressbar

while True:
    label_Space = Label(window, background = backgroundColor, foreground = foregroundColor, anchor = "center", text = "") #Creating the labels (MainConverter)
    label_Space.grid(column = 0, row = 0, columnspan = 3, sticky = "NSEW")
    label_Space.columnconfigure(1, weight = 1)
    label_Space.rowconfigure(1, weight = 1)

    label_Filename = Label(window, background = backgroundColor, foreground = foregroundColor, anchor ="center", text = "Enter file name: ")
    label_Filename.grid(column = 0, row = 2, sticky = "NSEW" )
    label_Filename.columnconfigure(1, weight = 1)
    label_Filename.rowconfigure(1, weight = 1)
    label_FilenameEntered = Label(window, background = backgroundColor, foreground = foregroundColor, anchor ="center", text = "Enter new file name: ")
    label_FilenameEntered.grid(column = 0, row = 3, sticky = "NSEW" )
    label_FilenameEntered.columnconfigure(1, weight = 1)
    label_FilenameEntered.rowconfigure(1, weight = 1)

    entry_Filename = Entry(window, background = backgroundColor, foreground = foregroundColor, textvariable = filename, width = 25, relief = "raised")
    entry_Filename.grid(column = 1, row = 2, sticky = "NSEW" )
    entry_Filename.columnconfigure(1, weight = 1)
    entry_Filename.rowconfigure(1, weight = 1)
    entry_FilenameNew = Entry(window, background = backgroundColor, foreground = foregroundColor, textvariable = filenameEntered, width = 25, relief = "raised")
    entry_FilenameNew.grid(column = 1, row = 3, sticky = "NSEW" )
    entry_FilenameNew.columnconfigure(1, weight = 1)
    entry_FilenameNew.rowconfigure(1, weight = 1)

    button_Convert = Button(window, background = backgroundColor, foreground = foregroundColor, anchor ="center", text = "Convert", command = clicked)#Creating the buttons (MainConverter)
    button_Convert.grid(column = 1, row = 6, sticky = "NSEW" )
    button_Convert.columnconfigure(1, weight = 1)
    button_Convert.rowconfigure(1, weight = 1)
    button_Reset = Button(window, background = backgroundColor, foreground = foregroundColor, anchor ="center", text = "Reset", command = reset)
    button_Reset.grid(column = 1, row = 7, sticky = "NSEW" )
    button_Reset.columnconfigure(1, weight = 1)
    button_Reset.rowconfigure(1, weight = 1)
    try:
        folder_icon = PhotoImage(file = "folder_icon.png")
        button_Open = Button(window, background = backgroundColor, foreground = foregroundColor, anchor ="center", image = folder_icon, command = selectFile)
    except:
        button_Open = Button(window, background = backgroundColor, foreground = foregroundColor, anchor ="center", text = "Select file", command = selectFile)
    button_Open.grid(column = 1, row = 1, sticky = "NSEW" )
    button_Open.columnconfigure(1, weight = 1)
    button_Open.rowconfigure(1, weight = 1)

    fromFiletypeChoose = ttk.Combobox(window, textvariable = fromFiletype, state = "readonly")#Creating the filetype select menues (MainConverter)
    fromFiletypeChoose["values"] = standartFiletypes
    fromFiletypeChoose.grid(column = 2, row = 2, sticky = "NSEW" )
    fromFiletypeChoose.columnconfigure(2, weight = 3)
    fromFiletypeChoose.rowconfigure(1, weight = 1)
    fromFiletypeChoose.current(0)

    toFiletypes = ("","")

    if fromFiletypeChoose.current() == 1:
        toFiletypes = (".txt",".xlsx")
    elif fromFiletypeChoose.current() == 2 or fromFiletypeChoose.current() == 6 or fromFiletypeChoose == 7:
        toFiletypes = ("")
    elif fromFiletypeChoose.current() == 3:
        toFiletypes = (".jpg",".tiff")
    elif fromFiletypeChoose.current() == 4:
        toFiletypes = (".csv",".txt",".xlsx")
    elif fromFiletypeChoose.current() == 5:
        toFiletypes = (".tiff")
    else:
        toFiletypes = standartFiletypes
        

    toFiletypeChoose = ttk.Combobox(window, textvariable = toFiletype, state = "readonly")
    toFiletypeChoose["values"] = toFiletypes
    toFiletypeChoose.grid(column = 2, row = 3, sticky = "NSEW" )
    toFiletypeChoose.columnconfigure(1, weight = 1)
    toFiletypeChoose.rowconfigure(1, weight = 1)
    toFiletypeChoose.current(0)

    menu_bar = Menu(window)#Creating the menubar
    menu_bar.config(background = backgroundColor)
    file_menu = Menu(menu_bar, tearoff = 0)
    file_menu.add_command(label="Settings", command = Settings, background = backgroundColor, foreground = foregroundColor)
    file_menu.add_command(label="Quit", command = Quit, background = backgroundColor, foreground = foregroundColor)

    help_menu = Menu(menu_bar, tearoff = 0)
    help_menu.add_command(label="Help", command = Help, background = backgroundColor, foreground = foregroundColor)

    menu_bar.add_cascade(label="File", menu = file_menu)
    menu_bar.add_cascade(label="Help", menu = help_menu)
    window.config(menu = menu_bar)

#reset()

    window.mainloop()

#Program for Det 157 cadet PIF digitization
#Created by C/Wolfe Spring 2023

#Import the required Libraries
from tkinter import *
from tkinter import ttk
import tkinter as tk
import openpyxl
from datetime import datetime
from PIL import Image,ImageTk
from pdf2image import convert_from_path
from PyPDF2 import PdfWriter, PdfReader
import os

#Create an instance of Tkinter frame
win= Tk()
labelsframe= tk.Frame(win,width=20) #Left frame for labels
labelsframe.pack(fill=tk.Y, side=tk.LEFT)
inputframe= tk.Frame(win,width=40) #Center frame for input
inputframe.pack(fill=tk.Y, side=tk.LEFT)
answersframe= tk.Frame(win,width=40) #Right frame for answers
answersframe.pack(fill=tk.Y, side=tk.LEFT)
picturesframe= tk.Frame(win,width=600, bg="light grey") #Right frame for answers
picturesframe.pack(fill=tk.Y, side=tk.RIGHT)

#Set the geometry of Tkinter frame
win.geometry("1100x750")
#Create a Canvas for Image
canvas = Canvas(picturesframe,width=600,height=750, bg="light grey")
canvas.pack()
#Get PDF
pdffilename = "example.pdf" #used to set file name
# Store Pdf pages as images with convert_from_path function
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print(str(current_time)+' Compiling images from '+str(pdffilename))
images = convert_from_path(pdffilename)#,poppler_path=r"D:\PythonD\Poppler\Release-23.01.0-0\poppler-23.01.0\Library\bin")
#Get number of pages in PDF
inputpdf = PdfReader(open(pdffilename, "rb"))
totalpages = len(inputpdf.pages)
pagenum = 0 #setting current page number
count=pagenum
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print(str(current_time)+' Populating images')
while (count<totalpages) : # Save pages as images of the pdf
    images[count].save('temporarypage'+ str(count) +'.jpg', 'JPEG')
    count = count+1
#Load image
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print(str(current_time)+' Population complete')
displaypage = pagenum
img= Image.open('temporarypage'+str(displaypage)+'.jpg')
#Resize the Image using resize method
resized_image= img.resize((514,725),Image.Resampling.LANCZOS)
new_image= ImageTk.PhotoImage(resized_image)
#Add image to Canvas
canvas.create_image(40,12,anchor=NW,image=new_image)
#Check for Excel Workbook with same name
excelfilecheck = os.path.isfile('FA_data.xlsx')
if excelfilecheck == False: #create new workbook if DNE
    wb = openpyxl.Workbook()
    wb.save('FA_data.xlsx')
#Create Class Folders for saving
if not os.path.exists(r"AS_100s"):
    os.makedirs(r"AS_100s")
if not os.path.exists(r"AS_200s"):
    os.makedirs(r"AS_200s")
if not os.path.exists(r"AS_300s"):
    os.makedirs(r"AS_300s")
if not os.path.exists(r"AS_400s"):
    os.makedirs(r"AS_400s")

#Create labels for labelsframe, inputframe, and answersframe 
labels = tk.Label(
   labelsframe,
   text="\nLast\nFirst\nAS Year\nFlight\nAge\nSex\n\nPush-ups\nSit-ups\
\nLap 1\nLap 2\nLap 3\nLap 4\nLap 5\nLap 6",
   justify= RIGHT,
   font=("Courier 13")
)
labels.pack()

space1 = tk.Label( #Used solely to crate spaces btwn certain entry boxes
   inputframe,
   text="",
   font=("Courier 9")
)
space2 = tk.Label( #Used solely to crate spaces btwn certain entry boxes
   inputframe,
   text="",
   font=("Courier 9")
)

conventionlabel = tk.Label(
   inputframe,
   text="",
   font=("Courier 11")
)

answers = tk.Label(
   answersframe,
   text="",
   justify= RIGHT,
   font=("Courier 13")
)
answers.pack()



#Function to update Answers (cued by button)
def display_answers():
   global var_lastname
   global var_firstname
   global var_asyear
   global var_flight
   global var_age
   global var_sex
   global var_push
   global var_sit
   global var_lap1
   global var_lap2
   global var_lap3
   global var_lap4
   global var_lap5
   global var_lap6
   
   global lastent #Last Name        Personal Info
   var_lastname= lastent.get()
   ansstring=var_lastname
   global firstent #First Name
   var_firstname= firstent.get()
   ansstring="\n"+ansstring+"\n"+var_firstname
   global asyearent #AS Year
   var_asyear= asyearent.get()
   ansstring=ansstring+"\nAS "+var_asyear
   global flightent #Flight
   var_flight= flightent.get()
   ansstring=ansstring+"\n"+var_flight
   global ageent #Age
   var_age= ageent.get()
   ansstring=ansstring+"\n"+var_age
   global sexent #Sex
   var_sex= sexent.get()
   ansstring=ansstring+"\n"+var_sex
   ansstring=ansstring+"\n"

   global pushent #Push-ups        Test Info
   var_push= pushent.get()
   global sitent #Sit-ups
   var_sit= sitent.get()
   global lap1ent #Lap 1
   var_lap1= lap1ent.get()
   global lap2ent #Lap 2
   var_lap2= lap2ent.get()
   global lap3ent #Lap 3
   var_lap3= lap3ent.get()
   global lap4ent #Lap 4
   var_lap4= lap4ent.get()
   global lap5ent #Lap 5
   var_lap5= lap5ent.get()
   global lap6ent #Lap 6
   var_lap6= lap6ent.get()
   ansstring=ansstring+"\n"+var_push+"\n"+var_sit+"\n"+var_lap1+"\n"+var_lap2+"\n"+var_lap3\
+"\n"+var_lap4+"\n"+var_lap5+"\n"+var_lap6
   answers.configure(text=ansstring) #Update Label
   global conventionent #naming convention
   var_convention= conventionent.get()
   global outputpdfname
   if var_lastname == "":
      outputpdfname = ""
   else:
      outputpdfname = str(var_lastname)+str(var_convention)+str(var_asyear)+".pdf"
   conventionlabel.configure(text=outputpdfname)
   #Check for file with same name in that AS year directory
   if var_asyear == "100" or var_asyear == "150":
       outputpdfpath = "AS_100s/"
   elif var_asyear == "200" or var_asyear == "250" or var_asyear == "500":
       outputpdfpath = "AS_200s/"
   elif var_asyear == "300":
       outputpdfpath = "AS_300s/"
   elif var_asyear == "400":
       outputpdfpath = "AS_400s/"
   else:
       outputpdfpath = ""
   global outputpdfpathnname
   outputpdfpathnname = outputpdfpath+outputpdfname
   pdffilecheck = os.path.isfile(outputpdfpathnname)
   if pdffilecheck == True: #notify that there is identical filename
       top = Toplevel(win)
       top.geometry("600x300")
       top.configure(bg="light grey")
       top.title("Warning")
       Label(top, text= "WARNING\n\nA file with the same name already exists",font=("Courier 13"),fg="red",bg="light grey").pack(pady=80)


#Function to flip viewing page (cued by button)
def flip_page():
   global displaypage
   global pagenum
   global img
   global resized_image
   global new_image
   if displaypage == pagenum:
      displaypage = pagenum+1
   else:
      displaypage = pagenum
   #update picture
   img= Image.open('temporarypage'+str(displaypage)+'.jpg')
   resized_image= img.resize((514,725),Image.Resampling.LANCZOS)
   new_image= ImageTk.PhotoImage(resized_image)
   canvas.create_image(40,12,anchor=NW,image=new_image)
   #Move focus to lap time lines
   lap1ent.focus_set()
   #Update pageent to display display page
   pageent.delete(0,tk.END)
   pageent.insert(0,displaypage)

#Function to jump to specified page (cued by button)
def goto_page():
   global displaypage
   global pagenum
   global img
   global resized_image
   global new_image
   global pageent
   ##Delete old temp images
   #os.remove('temporarypage'+str(pagenum)+'.jpg')
   #os.remove('temporarypage'+str(pagenum+1)+'.jpg')
   ## Store Pdf pages as images with convert_from_path function
   #images = convert_from_path(pdffilename)
   displaypage = int(pageent.get()) #updating current page number
   #count=pagenum
   #while (count<pagenum+2) : # Save pages as images of the pdf
   #    images[count].save('temporarypage'+ str(count) +'.jpg', 'JPEG')
   #    count = count+1
   #Load image
   if (displaypage % 2) ==0:
       pagenum = displaypage
   else:
       pagenum = displaypage-1
   img= Image.open('temporarypage'+str(displaypage)+'.jpg')
   #update picture
   img= Image.open('temporarypage'+str(displaypage)+'.jpg')
   resized_image= img.resize((514,725),Image.Resampling.LANCZOS)
   new_image= ImageTk.PhotoImage(resized_image)
   canvas.create_image(40,12,anchor=NW,image=new_image)
   
   

#Function to submit page and reset (cued by button)
def submit_reset():
   global var_lastname
   global var_firstname
   global var_asyear
   global var_flight
   global var_age
   global var_sex
   global var_push
   global var_sit
   global var_lap1
   global var_lap2
   global var_lap3
   global var_lap4
   global var_lap5
   global var_lap6
   #Export to Excel
   wb = openpyxl.load_workbook('FA_data.xlsx')
   sheet = wb['Sheet']
   rowData = [var_lastname,var_firstname,var_asyear,var_flight,var_age,\
var_sex,var_push,var_sit,var_lap1,var_lap2,var_lap3,var_lap4,var_lap5,var_lap6]
   sheet.append(rowData)
   wb.save('FA_data.xlsx')
   #Export Specific PDF Pages as Separate PDF
   global pdffilename
   global pagenum
   global outputpdfname
   global outputpdfpathnname
   inputpdf = PdfReader(open(pdffilename, "rb"))
   output = PdfWriter()
   output.add_page(inputpdf.pages[pagenum])
   output.add_page(inputpdf.pages[pagenum+1])
   with open(outputpdfpathnname, "wb") as outputStream:
       output.write(outputStream)

   #clear entry boxes
   lastent.delete(0,tk.END)
   firstent.delete(0,tk.END)
   asyearent.delete(0,tk.END)
   flightent.delete(0,tk.END)
   ageent.delete(0,tk.END)
   sexent.delete(0,tk.END)
   pushent.delete(0,tk.END)
   sitent.delete(0,tk.END)
   lap1ent.delete(0,tk.END)
   lap2ent.delete(0,tk.END)
   lap3ent.delete(0,tk.END)
   lap4ent.delete(0,tk.END)
   lap5ent.delete(0,tk.END)
   lap6ent.delete(0,tk.END)
   #update answers display label
   display_answers()
   #update pictures
   global displaypage
   global img
   global resized_image
   global new_image
   ##Delete old temp images
   #os.remove('temporarypage'+str(pagenum)+'.jpg')
   #os.remove('temporarypage'+str(pagenum+1)+'.jpg')
   ## Store Pdf pages as images with convert_from_path function
   #images = convert_from_path(pdffilename)
   pagenum = pagenum+2 #updating current page number
   #count=pagenum
   #while (count<pagenum+2) : # Save pages as images of the pdf
   #    images[count].save('temporarypage'+ str(count) +'.jpg', 'JPEG')
   #    count = count+1
   #Load image
   displaypage = pagenum
   img= Image.open('temporarypage'+str(displaypage)+'.jpg')
   #update picture
   img= Image.open('temporarypage'+str(displaypage)+'.jpg')
   resized_image= img.resize((514,725),Image.Resampling.LANCZOS)
   new_image= ImageTk.PhotoImage(resized_image)
   canvas.create_image(40,12,anchor=NW,image=new_image)
   #focus to first line
   lastent.focus_set()
   #Update pageent to display display page
   pageent.delete(0,tk.END)
   pageent.insert(0,displaypage)

#Create an Entry widgets to accept User Input
sizefont=10 #entry box font size/height
entwidth=30 #entry box width
space1.pack() # Adds space
lastent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for Last       Personal info
lastent.focus_set()
lastent.pack()
firstent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for First
firstent.pack()
asyearent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for AS Year
asyearent.pack()
flightent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for Flight
flightent.pack()
ageent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for Age
ageent.pack()
sexent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for Sex
sexent.pack()
space2.pack() # Adds space
pushent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Push-ups       Test info
pushent.pack()
sitent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Sit-ups
sitent.pack()
lap1ent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Lap 1
lap1ent.pack()
lap2ent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Lap 2
lap2ent.pack()
lap3ent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Lap 3
lap3ent.pack()
lap4ent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Lap 4
lap4ent.pack()
lap5ent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Lap 5
lap5ent.pack()
lap6ent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Lap 6
lap6ent.pack()
pageent= Entry(answersframe, width= 4, font=('Courier', sizefont))


#Create Buttons
ttk.Button(inputframe, text= "Verify",width= 20, command= display_answers).pack(pady=20)
ttk.Button(answersframe, text= "Flip Scorecard  ↻",width= 20, command=flip_page).pack(side=tk.BOTTOM, pady=20)
ttk.Button(answersframe, text= "Go to",width= 7, command=goto_page).pack(side=tk.BOTTOM)

pageent.pack(side=tk.BOTTOM)
#Create Naming Convention Input
conventionent= Entry(inputframe, width= entwidth, font=('Courier', sizefont)) #Entry input for naming convention
conventionent.pack()
conventionlabel.pack()
ttk.Button(inputframe, text= "Submit",width= 20, command= submit_reset).pack(pady=20)

win.mainloop()







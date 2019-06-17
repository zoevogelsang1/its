from tkinter import * 
from tkinter import filedialog
import pandas as pd
import pyautogui, time

import pyperclip
from PIL import Image
from openpyxl import load_workbook

window = Tk()
window.title("autofill")
window.geometry('400x300')

def getExcel():
	global wb
	import_file_path = filedialog.askopenfilename()
	wb = load_workbook (import_file_path, data_only = True) 

browseButton_Excel = Button(window, text='Import Excel File', command=getExcel)
browseButton_Excel.grid(column=1, row=0)
browseButton_Excel.grid(column=1, row=0)

lblsh = Label(window, text="Sheet Name ")
lblsh.grid(column=0, row=1)
txtsh= Entry(window,width=10)
txtsh.grid(column=1, row=1)

lbl0 = Label(window, text="First Col ")
lbl0.grid(column=0, row=2)
txt0= Entry(window,width=10)
txt0.grid(column=1, row=2)

lbl1 = Label(window, text="Start Index")
lbl1.grid(column=0, row=3)
txt1= Entry(window,width=10)
txt1.grid(column=1, row=3)

lbl2 = Label(window, text="End Index")
lbl2.grid(column=0, row=4)
txt2= Entry(window,width=10)
txt2.grid(column=1, row=4)


lbl3 = Label(window, text="Second Col")
lbl3.grid(column=2, row=2)
txt3= Entry(window,width=10)
txt3.grid(column=3, row=2)

lbl4 = Label(window, text="Start Index")
lbl4.grid(column=2, row=3)
txt4= Entry(window,width=10)
txt4.grid(column=3, row=3)

lbl5 = Label(window, text="End Index")
lbl5.grid(column=2, row=4)
txt5= Entry(window,width=10)
txt5.grid(column=3, row=4)

message = Label(window, text="")
message.grid(column=0, row=6)


def displaymessage():
	message.configure(text = "Attempting to find SIS site. Make sure that the page is visible and text field is not already selected. ")

def donemessage():
	message.configure(text = "")

def clicked():
	global cola, starta, enda, colb, startb, endb
	cola = str(txt0.get())
	starta= int(txt1.get())
	enda= int(txt2.get())
	colb = str(txt3.get())
	startb= int(txt4.get())
	endb= int(txt5.get())

	sheetname = txtsh.get()
	sheet = wb[sheetname]


	print("Inputs:", cola, starta, "-", enda, " ", colb, startb, "-", endb )

	datafield = Image.open('datafield.png')
	if datafield.mode == 'RGBA':
		datafield = datafield.convert('RGB')
	
	displaymessage()	
	center = None 
	while True: 
		im = pyautogui.screenshot()
		im.save("screenshot.png")
		print ("Attempting to find SIS site. Make sure that the page is visible and text field is not already selected. ")
		if im.mode == 'RGBA':
			im = im.convert('RGB')
	
		center = pyautogui.locateOnScreen(datafield)
		if center != None:
			donemessage()
			break;
	center = pyautogui.center(center)
	center = ( (center[0])/2, (center[1])/2.0 )
	pyautogui.click( center ) # text box should be clicked on 

	i = 0
	while int(starta)+i <= int(enda):
		pyautogui.click()
		cpy = sheet[cola + str(starta +i)].value
		pyperclip.copy(cpy)
		pyautogui.hotkey('command', 'v')		#paste 
		pyautogui.move(360,0)     			#move to transform
		pyautogui.click()					#click
		cpy = sheet[ colb +str(startb + i)].value	
		pyperclip.copy(cpy)	
		pyautogui.hotkey('command', 'v')		#paste
		pyautogui.move(105,0)				#moves to update. 
		pyautogui.click()
		pyautogui.click()
		if i >= 5:
			pyautogui.move(115,0)
			pyautogui.scroll(-15)
		else:
			pyautogui.move(115,38)
		pyautogui.click()
		pyautogui.move(-580, 0)
		i = i+1




enterbtn = Button(window, text="Enter", command=clicked)
enterbtn.grid(column=1, row=5)
enterbtn.grid(column=1, row=5)



window.mainloop()

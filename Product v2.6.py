#Developed by Neel Mishra, Mishraneel99@gmail.com .
#Created at 9/17/2017
#Automatic certificate Modifing and Emailing program.
#Created for Silver Oak Group of Institute.


####NOTE#############################################################################################################################################
#This is version 1 of my script.
#
#NOTE:
#The email credentials are not stored as raw string, but are taken as temporary string. And are automatically removed when the program terminates.
#
#THE REQUIREMENTS ARE......
#
#1) The excel file must be named as test.xlsx
#2) The designed image must be named as design.jpg
#
#WHAT's COMING UP IN THE NEXT VERSION
#
#1) GUI IMPLEMENTATION OF DIFFRENT OPTIONS
#2) Diffrent options, 
#	i) Option to save individual's pdf.
#	ii) Ability to select text location, and define the text size. This will genralize this program to work among any certificate design.
#3) User can have GUI path interface so that does not require to type a specific name of excel file. Instead he can browse directory with the GUI.
#4) Users can have same path interface with the certificate design so the process becomes easy.
#5) More implementations, depending on current limitations.
#                                                                           #############################################################################
from PyPDF2 import PdfFileMerger, PdfFileReader
import time
import PIL
from PIL import Image,ImageFont,ImageDraw,ImageTk
import xlrd
import PyPDF2
import os
from email import encoders
from email.mime.base import MIMEBase
import argparse
import smtplib
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import tkinter as Tk
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfile
from tkinter.messagebox import showerror
import random

#Defining globals
email = '' #The email from which the user will send certificates to the participants
serial = 170906001 # This is the serial, user can edit this to produce the required series.
random_password_string = '' # This will be the password participants will get in their mail
psw = 0 #The condition, to determine wether to encript the pdf files.
password = '' #The email password, the user will use.
condition = 0 #This variable will ensure proper program flow, depending on the user requirements
location = (800,735) #This is the coordinates of name of participants, which will be used to print name on certificates.
bottom = (120,1210) #This is the coordinates of serial.
epath = '' #This is the path of the certificate design. This will be fed from GUI
ipath = '' #This is the path of excel file, which will be fed from GUI.
temp = '' #Temporary folder path.
font_size = 120 #The font size which the user will feed.
serial_size = 40 #The serial font size, which the user will feed.
store_folder = '' #The folder in which all ther certificates will be stored.
pdf_files = []  #This will contain the information about the pdf names.

##Defining sub gui functions

def withdraw(event):
    global win,win3
    win3.destroy()
    win.deiconify()

def set_password(input_file, user_pass, owner_pass):
    """
    Function creates new temporary pdf file with same content,
    assigns given password to pdf and rename it with original file.
    """
    # temporary output file with name same as input file but prepended
    # by "temp_", inside same direcory as input file.
    path, filename = os.path.split(input_file)
    output_file = os.path.join(path,'protected_' + filename)
 
    output = PyPDF2.PdfFileWriter()
 
    input_stream = PyPDF2.PdfFileReader(open(input_file, "rb"))
 
    for i in range(0, input_stream.getNumPages()):
        output.addPage(input_stream.getPage(i))
 
    outputStream = open(output_file, "wb")
 
    # Set user and owner password to pdf file
    output.encrypt(user_pass, owner_pass, use_128bit=True)
    output.write(outputStream)
    outputStream.close()

def prev_label():
    pass

def bottom_shift_down(event):
    global win3,frame,bottom
    x,y = bottom
    y = y - 15
    bottom = x,y
    win3.destroy()
    set_serial("<Button-1>")

def bottom_shift_top(event):
    global win3,frame,bottom
    x,y = bottom
    y = y + 15
    bottom = x,y
    win3.destroy()
    set_serial("<Button-1>")

def bottom_shift_right(event):
    global win3,frame,bottom
    x,y = bottom
    x = x + 15
    bottom = x,y
    win3.destroy()
    set_serial("<Button-1>")

def bottom_shift_left(event):
    global win3,frame,bottom
    x,y = bottom
    x = x - 15
    bottom = x,y
    win3.destroy()
    set_serial("<Button-1>")

def location_shift_down(event):
    global win3,frame,location
    x,y = location
    y = y - 15
    location = x,y
    win3.destroy()
    set_font("<Button-1>")

def location_shift_top(event):
    global win3,frame,location
    x,y = location
    y = y + 15
    location = x,y
    win3.destroy()
    set_font("<Button-1>")

def location_shift_right(event):
    global win3,frame,location
    x,y = location
    x = x + 15
    location = x,y
    win3.destroy()
    set_font("<Button-1>")

def location_shift_left(event):
    global win3,frame,location
    x,y = location
    x = x - 15
    location = x,y
    win3.destroy()
    set_font("<Button-1>")

def pass_checkbox(event):
    global psw
    if psw == 0:
        psw = 1
    else:
        psw = 0

def random_password():
    global random_password_string
    random_password_string = chr(random.randrange(65,91)) + str(random.randint(100,1000)) + chr(random.randrange(97,123)) + chr(random.randrange(97,123)) + str(random.randint(100,1000))

def new_win(event):
    global window
    window.destroy()
    

def prev_op(event):
    global location,bottom,font_size,serial_size,ipath,window

    if ipath == '':
        print("Please set the image path!")
        return



    if location == None:
        print("Please set the location of text")
        return

    if bottom == None:
        print("Please set the location of serial")
        return
        
    image = Image.open(ipath.name)
    
    draw = ImageDraw.Draw(image)
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',font_size)
    serial_font = ImageFont.truetype('Borg.ttf',serial_size)
    #Drawing Sample Text
    draw.text(location, "Test Text" ,(0,0,0),font=font)
    #Drawing Sample Serial
    draw.text(bottom, "00230120" ,(0,0,0),font=serial_font)

    image.save('test.png')
    #win.withdraw()

    window = Tk.Toplevel()

    frame = Tk.Frame(window)
    frame.config(width = 1600, height = 1600)
    frame.grid()

    warning_label = Tk.Label(frame, text = "Please click on the image to exit",bg = "red", fg = "white")
    warning_label.config(font = ("Courier",32))
    warning_label.pack(fill = Tk.X)
    
    image2 = Tk.PhotoImage(file = 'test.png')
    label = Tk.Label(frame,image = image2)
    image2.image = image2
    label.pack()
    label.bind("<Button-1>",new_win)

def set_serial(event):
    global win3,bottom,serial_size
    """Menu 1 set font button event"""
    global ipath,frame

    if ipath == '':
        print("Please set the path of image")
        return
    win.withdraw()
    
    if bottom == None:
        bottom = (0,0)

    win3 = Tk.Toplevel()

    frame = Tk.Frame(win3)
    frame.pack()

    image = Image.open(ipath.name)
    
    draw = ImageDraw.Draw(image)
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',serial_size)
    #Drawing Sample Text
    draw.text(bottom, "055151524" ,(0,0,0),font=font)

    image.save('test.png')
    #win.withdraw()
    
    image2 = Tk.PhotoImage(file = 'test.png')
    label = Tk.Label(frame,image = image2)
    image2.image = image2
    label.pack()
    
    label.bind("<Button-1>",serial_func)
    win3.bind("<Left>",bottom_shift_left)
    win3.bind("<Down>",bottom_shift_top)
    win3.bind("<Up>",bottom_shift_down)
    win3.bind("<Right>",bottom_shift_right)
    win3.bind("<Button-3>",withdraw)
    win3.mainloop()

def pserial(image):
    global serial_size,bottom,serial
    draw = ImageDraw.Draw(image)
    size = serial_size
    
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',size)

    #Drawing name
    draw.text(bottom, str(serial) ,(0,0,0),font=font)
    int(serial)

def browse(event):
    '''Event function of browse of second radio option after cliking submit'''
    global store_folder,root2
    store_folder = askdirectory()
    print(store_folder)

    pass

def entry_one(event):
    global store_folder
    if store_folder == '':
        print("Please set a storage Location!")
        return
    op2()

def entry_many(event):
    global pdf_files
    if store_folder == '':
        print("Please set a storage Location!")
        return
    op3()
    

def serial_func(event):
    global win3
    global  bottom
    bottom = event.x,event.y
    print(bottom)
    win3.destroy()
    win.deiconify()
    return

def image_func(event):
    '''This is event function of imaging clicking event'''
    global win3
    global  location
    location = event.x,event.y
    print(location)
    win3.destroy()
    win.deiconify()
    return

def entry_submit_event(event):
    '''Menu1 First Submit button event'''
    global condition, epath, ipath, location,root2

    if epath == '' or ipath == '':
        print("Please specify the excel, and image path")
        return
    if location == None:
        print("Please sent a location of text, to modify the certificate")
        return

    if condition != 3 and psw == 1:
        print("Password feature, only works with email.")
        return
    
    if condition == 1:
        #Without email save all certificiate in one folder
        win.destroy()
        
        root2 = Tk.Tk()

        #Creating Frame
        
        frame = Tk.Frame(root2)
        frame.pack()

        #Creating label
        
        label = Tk.Label(frame, text = "Please choose the directory to save, by clicking the below button.",fg = "white", bg = "black", font = ("Courier",24))
        label.pack(fill = Tk.X)

        #Creating Set Location button
        
        button = Tk.Button(frame, text = "Set Location",bg = "yellow")
        button.config(width = 13, font = ("Courier", 24))
        button.pack()
        button.bind("<Button-1>",browse)

        #Creating Submit label

        label = Tk.Label(frame, text = "After setting location, submit below.",fg = "white", bg = "black", font = ("Courier",24))
        label.pack(fill = Tk.X)

        #Creating Submit button
        
        entry_button = Tk.Button(frame, text = "Submit",bg = "Red",fg= "white")
        entry_button.config(width = 13, font = ("Courier", 24))
        entry_button.pack()
        entry_button.bind("<Button-1>",entry_one)

        root2.mainloop()
    elif condition == 2:
        #Without emailing save all certificate as one pdf
        win.destroy()
    
        root2 = Tk.Tk()

        #Creating Frame
        
        frame = Tk.Frame(root2)
        frame.pack()

        #Creating label
        
        label = Tk.Label(frame, text = "Please choose the directory to save, by clicking the below button.",fg = "white", bg = "black", font = ("Courier",24))
        label.pack(fill = Tk.X)

        #Creating Set Location button
        
        button = Tk.Button(frame, text = "Set Location",bg = "yellow")
        button.config(width = 13, font = ("Courier", 24))
        button.pack()
        button.bind("<Button-1>",browse)

        #Creating Submit label

        label = Tk.Label(frame, text = "After setting location, submit below.",fg = "white", bg = "black", font = ("Courier",24))
        label.pack(fill = Tk.X)

        #Creating Submit button
        
        entry_button = Tk.Button(frame, text = "Submit",bg = "Red",fg= "white")
        entry_button.config(width = 13, font = ("Courier", 24))
        entry_button.pack()
        entry_button.bind("<Button-1>",entry_many)

        root2.mainloop()

    elif condition == 3:
        #Mail  all the ceritificates without saving
        global email,password
        win.destroy()
        os.system('cls')
        email = input("Please enter an email\n")
        password = input("Please enter a password\n")
        try:
            os.system('cls')
        except:
            os.system('clear')
        op1()
        win2 = Tk.Tk()
        frame3 = Tk.Frame(win2)
        frame3.pack(side = Tk.TOP)
        info = Tk.Label(frame3, text = "Mailed, certificates to every participant. Successfully!", fg = 'violet', bg = 'white')
        info.configure(width = 200)
        info.config(font = ("Courier", 32))
        info.pack(fill = Tk.X)
        win2.mainloop()

    else:
        print("Please select an option, and try again")

def radio1_event(event):
    '''First meanu First radio button even'''
    global condition
    condition = 1
    
def radio2_event(event):
    '''First menu Second radio button event'''
    global condition
    condition = 2

def radio3_event(event):
    '''First menu Third radio button event'''
    global condition
    condition = 3

def browse_image(event):
    """Menu1 Set Image path button event"""
    global ipath
    win.withdraw()
    ipath = askopenfile(parent=win,mode='rb',title='Choose the certificate design')
    win.deiconify()
    return

def browse_excel(event):
    """Menu1 Set excel path button event"""
    global epath
    win.withdraw()
    epath = askopenfile(parent=win,mode='rb',title='Choose the Excel file')
    win.deiconify()
    return

def set_font(event):
    global win3,location,font_size
    """Menu 1 set font button event"""
    global ipath,frame

    if ipath == '':
        print("Please set the path of image")
        return
    win.withdraw()
    
    if location == None:
        location = (0,0)

    win3 = Tk.Toplevel()

    frame = Tk.Frame(win3)
    frame.pack()

    image = Image.open(ipath.name)
    
    draw = ImageDraw.Draw(image)
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',font_size)
    #Drawing Sample Text
    draw.text(location, "Test Text" ,(0,0,0),font=font)

    image.save('test.png')
    #win.withdraw()
    
    image2 = Tk.PhotoImage(file = 'test.png')
    label = Tk.Label(frame,image = image2)
    image2.image = image2
    label.pack()
    
    label.bind("<Button-1>",image_func)
    label.bind("<Button-3>",withdraw)
    win3.bind("<Left>",location_shift_left)
    win3.bind("<Down>",location_shift_top)
    win3.bind("<Up>",location_shift_down)
    win3.bind("<Right>",location_shift_right)
    win3.mainloop()
##End of defining sub gui functions

#GUI FUNCTION AKA MAIN FUNCTION
    
#Creating new window, and giving it a title
win = Tk.Tk()
win.resizable(1,0)

win.title("Product v 1.1")


#Creating entry frame
frame1 = Tk.Frame(win)
#Fixing the frame
frame1.pack()

#Creating Secondary frame
bottom_frame = Tk.Frame(win)

#Fixing the frame
bottom_frame.pack(side = Tk.BOTTOM)
#Creating an entry Label
entry_label = Tk.Label(frame1, text = "Welcome, please select one of the following options.",fg = "white",bg = "black")
entry_label.config(width=300)
entry_label.config(font=("Courier", 32))
entry_label.pack(fill = Tk.X)

#Creating, and configuring radio entry buttons
option1 = Tk.Radiobutton(frame1, text = "Without emailing, Save all certificate as pdf in one folder.",value = 1)
option1.config(width=200)
option1.config(font=("Courier", 28))
option2 = Tk.Radiobutton(frame1, text = "Without emailing, Save all certificate as an entire pdf.",value = 2)
option2.config(width=200)
option2.config(font=("Courier", 28))
option3 = Tk.Radiobutton(frame1, text = "Email all certificate, without saving. ",value = 3)
option3.config(width=200)
option3.config(font=("Courier", 28))

#Packing radio entry buttons
option1.pack(fill = Tk.X)
option2.pack(fill = Tk.X)
option3.pack(fill = Tk.X)

#Binding radio buttons
option1.bind("<Button-1>", radio1_event)
option2.bind("<Button-1>", radio2_event)
option3.bind("<Button-1>", radio3_event)

#Creating Label for path defining buttons

path_label = Tk.Label(frame1, text = "Configure the path of excel, and image file below!", bg = "black" , fg = "white", width = 200)
path_label.configure(font=("Courier", 26))
path_label.pack()

#Creating the path buttons

excel_button = Tk.Button(frame1, text = "Excel Path", bg = "Yellow", fg = "black")
excel_button.config(width = 12 ,font = ("Courier",28))
excel_button.bind("<Button-1>",browse_excel)
excel_button.pack(side = Tk.TOP)

image_button = Tk.Button(frame1, text = "Image Path", bg = "yellow", fg = "black")
image_button.config(width = 12, font = ("Courier", 28))
image_button.bind("<Button-1>",browse_image)
image_button.pack(side = Tk.TOP)

#Creating Label for Set Text Location.

text_label = Tk.Label(frame1, text = "Set the location of name, in certificate image below!", bg = "black" , fg = "white", width = 200)
text_label.configure(font=("Courier", 26))
text_label.pack()

#Creating Set Text button.

text_button = Tk.Button(frame1, text = "Set Text", bg = "Yellow", fg = "black")
text_button.config(width = 12 ,font = ("Courier",28))
text_button.bind("<Button-1>",set_font)
text_button.pack(side = Tk.TOP)

#Creating Set Serial button.
serial_button = Tk.Button(frame1, text = "Set Serial", bg = "Yellow", fg = "black")
serial_button.config(width = 12, font = ("Courier",28))
serial_button.bind("<Button-1>",set_serial)
serial_button.pack()

#Creating submit label
entry_label = Tk.Label(frame1, text = "Click the button to SUBMIT!", bg = "black" , fg = "white", width = 200)
entry_label.configure(font=("Courier", 26))
entry_label.pack()

#Creating password Checkbox button.
password_checkbox = Tk.Checkbutton(frame1, text =  "Password")
password_checkbox.configure(font = ("Courier", 20))
password_checkbox.pack()
password_checkbox.bind("<Button-1>",pass_checkbox)

#Creating Preview button
entry_preview = Tk.Button(frame1, text= "PREVIEW",bg = "blue", fg = "white")
entry_preview.config(width=8)
entry_preview.config(font=("Courier", 28))
entry_preview.pack()
entry_preview.bind("<Button-1>",prev_op)

#Creating submit button

entry_submit = Tk.Button(frame1, text= "SUBMIT",bg = "red", fg = "white")
entry_submit.config(width=6)
entry_submit.config(font=("Courier", 28))
entry_submit.pack()

#Binding entry submit button
entry_submit.bind("<Button-1>", entry_submit_event)


#Subfunctions

def convert_single(image):
    global random_password_string,psw
    if psw == 1:
        random_password()
    if image.mode == "RGBA":
        image.convert("RGB")
    image.save("Certificate.pdf","PDF",resolution = 100.0)
    if psw == 1:
        set_password("certificate.pdf",random_password_string,random_password_string)

def convert_many(image,name,sirname):
    global  store_folder
    if image.mode == "RGBA":
        image.convert("RGB")
    image.save(store_folder+'/' + name+ ' ' + sirname+".pdf","PDF",resolution = 100.0)

def convert_as_one(image,leng):
    global store_folder,pdf_files,temp
    i = 0
    image.save(temp + str(leng) +'.pdf')
    
    merger = PdfFileMerger()
    for files in pdf_files:
        i = i + 1
        merger.append(temp + files)
    merger.write(store_folder + '/' + "merged.pdf")
    print("Sucess")
    
def conver_as_one(image):
    pass
    
def fsize(full_name):
    '''This automatically defines font size'''
    if len(full_name) in range (15,20):
        size = font_size - 4
    elif len(full_name) in range(0,15):
        size = font_size - 2
    elif len(full_name) in range (20,25):
        size = font_size - 6
    else:
        size = font_size - 8
    return size


def send_mail(server,receiver, name, sirname):#Debugged

    ##Sending mail
    global email,password,random_password_string
    #Creating multipart object
    msg = MIMEMultipart()

    #Opening pdf
    #Converting into MIME both image and text
    if psw == 1:
        file = open('protected_certificate.pdf','rb')
        text = MIMEText('This is Testing Phase' + ' ' 'The password for opening of pdf is' + ' ' + random_password_string)
    else:
        file = open('certificate.pdf','rb')
        text = MIMEText('This is Testing Phase')
    
    

    #Adding values to msg dictionary..
    msg['To'] = name + ' ' + sirname
    msg['Subject'] = 'Testing subject'
    msg['From'] = 'Silver Oak Group of Institutes'

    #appending message , and image body to our message from cert(name,sirname):
    msg.attach(text)
    attach = MIMEBase('application', "octet-stream")
    attach.set_payload(file.read())
    encoders.encode_base64(attach)

    attach.add_header('Content-Disposition', 'attachment', filename='Certificate.pdf')

    msg.attach(attach)
    server.sendmail(email,receiver,msg.as_string())

    
def cert(name, sirname):  #Debugged
    'This function modifies the certificate'
    global ipath,location,font_size,pdf_files,temp
    #Set the height and width of the text field here.
    width,height = location

    #Set the path to store here.
    path = ''

    #Opening a image file
    filename = ipath.name
    new = Image.open(filename)

    if new.mode != "RGB":
        new = new.convert("RGB")

    #Creating a draw object
    draw = ImageDraw.Draw(new)
    #size = fsize(name + '' + sirname)
    size = font_size
    
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',size)

    #Drawing name
    draw.text(location,name + " " + sirname ,(0,0,0),font=font)
    pserial(new)
    #Converting the image into pdf
    convert_single(new)

def cert2(name,sirname):
    'This function modifies the certificate'
    global ipath,location,font_size
    #Set the height and width of the text field here.
    width,height = location

    #Set the path to store here.
    path = store_folder

    #Opening a image file
    filename = ipath.name
    new = Image.open(filename)

    if new.mode != "RGB":
        new = new.convert("RGB")

    #Creating a draw object
    draw = ImageDraw.Draw(new)
    #size = fsize(name + '' + sirname)
    size = font_size
    
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',size)

    #Drawing name
    draw.text(location,name + " " + sirname ,(0,0,0),font=font)
    pserial(new)
    #Converting the image into pdf
    convert_many(new,name,sirname)

def cert3(name,sirname):
    'This function modifies the certificate'
    global ipath,location,font_size
    #Set the height and width of the text field here.
    width,height = location

    #Set the path to store here.
    path = store_folder

    #Opening a image file
    filename = ipath.name
    new = Image.open(filename)

    if new.mode != "RGB":
        new = new.convert("RGB")

    #Creating a draw object
    draw = ImageDraw.Draw(new)
    #size = fsize(name + '' + sirname)
    size = font_size
    
    #Assinging a text font.
    font = ImageFont.truetype('Borg.ttf',size)

    #Drawing name
    draw.text(location,name + " " + sirname ,(0,0,0),font=font)
    pserial(new)
    #Converting the image into pdf
    convert_as_one(new,len(pdf_files))

def op1():#Debugged
    
    '''This function contacts server, and sends mail according to the excel file'''
    global email,password,epath,serial
    counter = 0
    #Enter the name of the Excel file.
    filename = epath.name

    #Opening workbook.
    workbook = xlrd.open_workbook(filename)

    #Opening worksheet.
    worksheet = workbook.sheet_by_index(0)
    
    ##Contacting server.

    #Creating server object.
    server = smtplib.SMTP('smtp.gmail.com', 587)

    #Saying hello to server , for making our presence felt.
    server.ehlo()

    #Encrypting further communication
    server.starttls()

    #Logging in
    server.login(email, password)

    #Sending every entry it's certificate.
    
    for row in range(worksheet.nrows):

        #comment the below if, if there are no description for attributes.
        if row == 0:
            continue

        #Initializing temporary variables,for passing in the cert function
        name = (worksheet.cell_value(row,0))
        email = (worksheet.cell_value(row,2))
        sirname = (worksheet.cell_value(row,1))

        #Passing the values to cert functions
        cert(name,sirname)
        serial = serial + 1
        #convert('C:\Users\Neel\Downloads\Test\Developingphase\design.jpg')
        send_mail(server,email,name,sirname)
        counter += 1
        print("Amount of mail sent =",counter)

    #Closing email
    server.close()

def op2():#Debugged
    
    '''This function contacts server, and sends mail according to the excel file'''
    global epath,serial
    #Enter the name of the Excel file.
    filename = epath.name

    #Opening workbook.
    workbook = xlrd.open_workbook(filename)

    #Opening worksheet.
    worksheet = workbook.sheet_by_index(0)

    #Creating every entry it's certificate.
    
    for row in range(worksheet.nrows):

        #comment the below if, if there are no description for attributes.
        if row == 0:
            continue

        #Initializing temporary variables,for passing in the cert function
        name = (worksheet.cell_value(row,0))
        sirname = (worksheet.cell_value(row,1))

        #Passing the values to cert functions
        cert2(name,sirname)
        serial = serial + 1

def op3():
    '''This function contacts server, and sends mail according to the excel file'''
    global epath,pdf_files,temp,serial
    #Enter the name of the Excel file.
    filename = epath.name

    #Opening workbook.
    workbook = xlrd.open_workbook(filename)

    #Opening worksheet.
    worksheet = workbook.sheet_by_index(0)
    #Creating every entry it's certificate.
    i = 0
    
    temp = store_folder + '/temp/'
    if not os.path.exists(temp):
            os.makedirs(temp)
    for row in range(worksheet.nrows):

        #comment the below if, if there are no description for attributes.
        if row == 0:
            continue

        #Initializing temporary variables,for passing in the cert function
        name = (worksheet.cell_value(row,0))
        sirname = (worksheet.cell_value(row,1))

        #comment the below if, if there are no description for attributes.
        i += 1
        pdf_files += [str(i) + '.pdf']
        cert3(name,sirname)
        serial = serial + 1
        
#main


def excel():

    #Enter the name of the Excel file.
    filename = 'Test.xlsx'

    #Opening workbook.
    workbook = xlrd.open_workbook(filename)

    #Opening worksheet.
    worksheet = workbook.sheet_by_index(0)



#Defining the loop of window
win.mainloop()

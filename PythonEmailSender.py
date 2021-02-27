# Programmer - python_scripts (Abhijith Warrier)

# PYTHON GUI TO SEND EMAIL WITH MULTIPLE ATTACHMENTS USING smtplib AND email MODULES

# Simple Mail Transfer Protocol (SMTP) is a protocol, which handles sending e-mail and routing e-mail between mail
# servers. The smtplib module defines an SMTP client session object that can be used to send mail to any Internet
# machine with an SMTP or ESMTP listener daemon.
#
# The email module is a library for managing email messages. It is specifically not designed to do any sending of email
# messages to SMTP, NNTP, or other servers; those are functions of modules such as smtplib and nntplib.
# The email package attempts to be as RFC-compliant as possible
#
# Enable Less secure app access to send mail from Gmail Account using this Python GUI. To enable follow the below steps
# Login to accounts.goole.com -> Click on Security -> Scroll down and Turn On the option called 'Less secure app access'

# Importing necessary packages
import os
import smtplib
import tkinter as tk
from tkinter import *
from email import encoders
from tkinter import messagebox
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter.filedialog import askopenfilenames

# Defining CreateWidgets() to create necessary tkinter widgets
def CreateWidgets():
    labelfromEmail = Label(root, text='EMAIL - ID : ', bg='darkslategray4', font=('', 15, 'bold'))
    labelfromEmail.grid(row=0, column=0, pady=5, padx=5)

    root.entryfromEmail = Entry(root, width=50, textvariable=fromEmail)
    root.entryfromEmail.grid(row=0, column=1, pady=5, padx=5)

    labelpasswordEmail = Label(root, text='PASSWORD : ', bg='darkslategray4', font=('', 15, 'bold'))
    labelpasswordEmail.grid(row=1, column=0, pady=5, padx=5)

    root.entrypasswordEmail = Entry(root, width=50, textvariable=passwordEmail, show='*')
    root.entrypasswordEmail.grid(row=1, column=1, pady=5, padx=5)

    root.showhideBTN = Button(root, text='SHOW', command=showPassword, width=10)
    root.showhideBTN.grid(row=1, column=2, pady=5, padx=5)

    labeltoEmail = Label(root, text='TO EMAIL - ID : ', bg='darkslategray4', font=('', 15, 'bold'))
    labeltoEmail.grid(row=2, column=0, pady=5, padx=5)

    root.entrytoEmail = Entry(root, width=50, textvariable=toEmail)
    root.entrytoEmail.grid(row=2, column=1, pady=5, padx=5)

    labelsubjectEmail = Label(root, text='SUBJECT : ', bg='darkslategray4', font=('', 15, 'bold'))
    labelsubjectEmail.grid(row=3, column=0, pady=5, padx=5)

    root.entry_subjectEmail = Entry(root, width=50, textvariable=subjectEmail)
    root.entry_subjectEmail.grid(row=3, column=1, pady=5, padx=5)

    labelattachmentEmail = Label(root, text='ATTACHMENT : ', bg='darkslategray4', font=('', 15, 'bold'))
    labelattachmentEmail.grid(row=4, column=0, pady=5, padx=5)

    root.entryattachmentEmail = Text(root, width=65, height=5)
    root.entryattachmentEmail.grid(row=4, column=1, pady=5, padx=5)

    attachmentBTN = Button(root, text='BROWSE', command=fileBrowse, width=10)
    attachmentBTN.grid(row=4, column=2, pady=5, padx=5)

    labelbodyEmail = Label(root, text='MESSAGE : ', bg='darkslategray4', font=('', 15, 'bold'))
    labelbodyEmail.grid(row=5, column=0)

    root.bodyEmail = Text(root, width=100, height=20)
    root.bodyEmail.grid(row=6, column=0, columnspan=3, pady=5, padx=5)

    sendEmailBTN = Button(root, text='SEND EMAIL', command=sendEmail, width=10)
    sendEmailBTN.grid(row=7, column=2, padx=5, pady=5)

    exitBTN = Button(root, text='EXIT', command=emailExit, width=10)
    exitBTN.grid(row=7, column=0, padx=5, pady=5)

# Defining the showPassword() to show the password instead of the masking
def showPassword():
    # Configuring the button to show the text as HIDE and run hidePassword() when clicked
    root.showhideBTN.config(text='HIDE', command=hidePassword)
    # Setting the show attribute to empty string to show the password
    root.entrypasswordEmail.config(show='')

# Defining the hidePassword() to mask the password
def hidePassword():
    # Configuring the button to show the text as HIDE and run showPassword() when clicked
    root.showhideBTN.config(text='SHOW', command=showPassword)
    # Setting the show attribute to * to maske the password
    root.entrypasswordEmail.config(show='*')

# Defining fileBrowse() to browse and select files which are supposed to be sent as attachments
def fileBrowse():
    # Presenting the user with file selection dialog to select the files that are to be sent.
    # askopenfilenames() function can be used to select multiple files
    # Setting initialdir is optional
    root.filename = askopenfilenames(initialdir='YOUR DIRECTORY PATH')
    # Looping thorugh the selected files and displaying them in attacmentEntry widget
    for files in root.filename:
        # Fetching only the file names from the path using the os.path.basename() method
        filename = os.path.basename(files)
        root.entryattachmentEmail.insert('1.0', filename + '\n')

# Defining emailExit() to exit from the GUI
def emailExit():
    # Getting the confirmation from the user using the messagebox
    MsgBox = messagebox.askquestion('Exit Application', 'Are you sure you want to exit?')
    if MsgBox == 'yes':
        # Closing the GUI window
        root.destroy()

# Defining sendEmail() to send the email
def sendEmail():
    # Fetching all the user-input parameters and storing in respective variables
    fromEmail1 = fromEmail.get()
    passwordEmail1 = passwordEmail.get()
    toEmail1 = toEmail.get()
    subjectEmail1 = subjectEmail.get()
    bodyEmail1 = root.bodyEmail.get('1.0', END)

    # Creating instance of class MIMEMultipart()
    message = MIMEMultipart()
    # Storing the email details in respective fields
    message['From'] = fromEmail1
    message['To'] = toEmail1
    message['Subject'] = subjectEmail1
    # Attach message with MIME instance
    message.attach(MIMEText(bodyEmail1))

    # Iterating through the files in attachment list
    for files in root.filename:
        # Opening and reading the file into attachment
        attachment = open(files, 'rb').read()
        # Creating instance of MIMEBase and naming it as emailAttach
        emailAttach = MIMEBase('application', 'octet-stream')
        # Changing the payload into encoded form
        emailAttach.set_payload(attachment)
        # Encoding the attachment into base 64
        encoders.encode_base64(emailAttach)
        # Adding headers to the files
        emailAttach.add_header('Content-Disposition', 'attachment; filename= %s' % os.path.basename(files))
        # Attaching the instane emailAttach to the message instance
        message.attach(emailAttach)

    # Sending the email with attachments
    try:
        # Create a smtp session
        smtp = smtplib.SMTP('smtp.gmail.com', 587)
        # Starting TLS for security
        smtp.starttls()
        # Authenticate the user
        smtp.login(fromEmail1, passwordEmail1)
        # Sending the email with Mulitpart message converted into string
        smtp.sendmail(fromEmail1, toEmail1, message.as_string())
        messagebox.showinfo('SUCCESS', 'EMAIL SENT TO ' + str(toEmail1))
        # Terminating the session
        logout = smtp.quit()

    # Catching authenctication error
    except smtplib.SMTPAuthenticationError:
        messagebox.showerror('ERROR', 'INVALID USERNAME OR PASSWORD')
    # Catching connection error
    except smtplib.SMTPConnectError:
        messagebox.showerror('ERROR', 'PLEASE TRY AGAIN LATER')

# Creating object of tk class
root = tk.Tk()

# Setting the title and background color disabling the resizing property
root.title('PythonMailer')
root.config(background='darkslategray4')
root.resizable(False, False)
root.geometry('720x580')

# Creating tkinter variables
toEmail = StringVar()
fromEmail = StringVar()
passwordEmail = StringVar()
subjectEmail = StringVar()

# Calling the CreateWidgets() function with argument bgColor
CreateWidgets()

# Defining infinite loop to run application
root.mainloop()
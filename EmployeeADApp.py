"""
Authored by: Adam Hollow - DevOps Technician, Devices.
For: South Ayrshire Council
"""

import csv
import ctypes
import sys
import datetime
from tkinter import *
from tkinter import font
from tkinter.ttk import *
import win32con
import adConnect
import encryptIT
import xmlParse

filePath = r'\\PATH_TO_STORAGE_FOLDER\inf_' + adConnect.full_name + '.xml'  # Create filepath var
adConnect.pyad.pyad_setdefaults(ldap_server="LDAPSERVER", username="SERVICE_ACCNT",
                                password=encryptIT.decrKey())
MessageBox = ctypes.windll.user32.MessageBoxW  # Instantiation of message box alias
app = Tk()  # Instantiation of app window
helper = Tk()  # Instantiation of help window
helper.geometry("0x0")  # Help dimensions
helper.overrideredirect(True)  # Setting help to hidden
appFontNORM = font.Font(family='Helvetica', size=8)  # App font normal
appFontBOLD = font.Font(family='Helvetica', size=10, weight='bold')  # App font bold

user = adConnect.userCon()  # Set user variable
update = user.update_attribute  # set function variable

'''
FUNCTION DESCRIPTION
change_ATTRIBUTE():
        Update user via PyAD's update_attribute().
'''


def change_Attr(AD_ATT, APP_ATT):  # alias to update attributes
    update(AD_ATT, APP_ATT)


def gain_defaults():
    app.protocol('WM_DELETE_WINDOW', lambda: incrementCount())  # Sets close button to incrementCount()

    # Grabs values currently already entered in to AD and formats them for display.
    # OBTAINING
    locationCurrent = str(user.get_attribute("physicalDeliveryOfficeName"))
    empID = str(user.get_attribute("EmployeeID"))
    etarmisLogin = user.get_attribute("etarmisNumber")
    badgeID = user.get_attribute("badgeID")
    # FORMATTING AND INSERTION
    employee_id_input.insert(2, empID[2:-2])
    etarmis_login_input.insert(2, etarmisLogin)
    flexi_badge_input.insert(2, badgeID)
    try:
        '''due to location being a combo box, the formatting or value of this in AD prior to running the app
           may not match up with a value in the list. Encase in try;except block in order to ignore if there
           are any errors.'''
        location_input.current(
            location.index(locationCurrent[2:-2]))  # Attempt to find location in list, without quotes
    except:
        pass


def gen_CSV():
    # Generation of CSV file holding data.
    SAM_AN = str(user.get_attribute("SAMAccountName"))[2:-2]  # Account Login ID
    EmployeeID = employee_id_input.get()  # Employee ID or Number
    Location = location_input.get()  # Primary Location
    namelist = adConnect.full_name.split(", ")  # Split full name into first/last
    eLogin = etarmis_login_input.get()  # Etarmis Login
    fBadge = flexi_badge_input.get()  # Badge Number
    pEmail = personal_email_input.get()  # Personal E-Mail
    pPhone = personal_phone_input.get()  # Personal Phone
    pPhoneType = ""
    captureDate = datetime.datetime.now().strftime('%d/%m/%Y')

    if str(pPhone[0:2]) == "01":
        pPhoneType = "Home"
    elif str(pPhone[0:2]) == "07":
        pPhoneType = "Mobile"
    try:
        # Inside try block for less common names
        fName = namelist[1]  # First name from namelist array in [1] position
        lName = namelist[0]  # Last name from namelist array in [0] position
    except:
        namelist = adConnect.full_name.split(" ")  # If split with comma failed, split on space
        fName = namelist[0]  # First name from namelist array in [0] position
        lName = namelist[1]  # Last name from namelist array in [1] position

    pPhone = pPhone.replace(" ", "")
    firstpart, secondpart = pPhone[:len(pPhone) // 2], pPhone[len(pPhone) // 2:]
    pPhone = firstpart + " " + secondpart
    # Open CSV file
    with open(r'\\Path_to_storage_folder\Emp_Info.csv', 'a', newline='') as f:
        writer = csv.writer(f)  # Instantiate writer as writer
        # Write below variables to csv in specified order
        writer.writerow([encryptIT.f.encrypt(str.encode(SAM_AN)), encryptIT.f.encrypt(str.encode(EmployeeID)),
                         encryptIT.f.encrypt(str.encode(fName)), encryptIT.f.encrypt(str.encode(lName)),
                         encryptIT.f.encrypt(str.encode(Location)), encryptIT.f.encrypt(str.encode(eLogin)),
                         encryptIT.f.encrypt(str.encode(fBadge)), encryptIT.f.encrypt(str.encode(pPhone)), encryptIT.f.encrypt(str.encode(pPhoneType)),
                         encryptIT.f.encrypt(str.encode(pEmail)), encryptIT.f.encrypt(str.encode(captureDate))])


# Function called from button clicked.
def callback():
    #EMPLOYEE ID VALIDATION
    # Make sure employee id is at least 6 chars long, if not 7 chars in length then add leading 0's
    if employee_id_input.get() == "" or location_input.get() == "" \
            or flexi_badge_input.get() == "":
        # Messagebox to advise required fields are not filled out
        MessageBox(None, "Please make sure all fields marked with an asterisk(*) are filled.", "Entry Error", 0)
        return 0
    try:
        z = int(employee_id_input.get())
    except ValueError:
        # Messagebox to advise employee id must be numbers
        MessageBox(None, "Your employee ID must only contain numbers, do not include 'SAC'.", "Value Error", 0)
        return 0
    if len(employee_id_input.get()) < 5:
        # Messagebox to advise required fields are not filled out
        MessageBox(None, "Employee ID field is mandatory, please make sure you have filled this in with the correct "
                         "information.", "Employee ID", 0)
        return 0
    empID = str(employee_id_input.get())
    length = 7 - len(empID)
    if length > 0:
        zeros = "0" * length
        empID = zeros + empID

    # PERSONAL E-MAIL VALIDATION
    if str(personal_email_input.get()) != "":
        try:
            x = str(personal_email_input.get()).index("@")
            if str(personal_email_input.get())[(x+1):].lower() == "south-ayrshire.gov.uk":
                # Messagebox to advise internal e-mails not allowed.
                MessageBox(None, "@south-ayrshire.gov.uk is not valid for a personal e-mail.", "Entry Error", 0)
                return 0
        except ValueError:
            # Messagebox to advise email invalid
            MessageBox(None, "Invalid E-mail. Must be formatted 'example@domain.com'.", "Invalid E-mail", 0)
            return 0

    # FLEXI BADGE VALIDATION
    if len(flexi_badge_input.get()) < 6:
        # Messagebox to advise Badge Number not long enough
        MessageBox(None, "Please make sure your ID badge number is at least 6 numbers in length.", "Entry Error", 0)
        return 0

    # PERSONAL PHONE VALIDATION
    try:
        if personal_phone_input.get() != "":
            s = int(personal_phone_input.get())
    except ValueError:
        # Messagebox to advise phone number is not a number
        MessageBox(None, "Personal phone must only contain numbers. Please do not include spaces.", "Value Error", 0)
        return 0
    if len(personal_phone_input.get()) < 11 or len(personal_phone_input.get()) > 12:
        if personal_phone_input.get() == "":
            # Messagebox to advise phone number is blank.
            MessageBox(None, "Personal phone is a required field.", "Value Error", 0)
            return 0
        # Messagebox to advise Phone number is not valid
        MessageBox(None, "Personal phone must be between 11 and 12, inclusive, characters in length.", "Value Error", 0)
        return 0
    elif str(personal_phone_input.get())[0:2] != '01':
        if str(personal_phone_input.get())[0:2] != '07':
            # Messagebox to advise Phone number is not valid
            MessageBox(None, "Personal phone must start with a 01 or 07.", 0)
            return 0

    # Yes/No box asking to confirm details
    confirm = MessageBox(None,
                         "Please confirm the details you have entered are correct. If you are sure, click 'Yes'.",
                         "Confirm Details", win32con.MB_YESNO)  # Asking user to confirm details are correct.
    if confirm == 6:  # In case of 'Yes' proceed to change details.
        try:
            # Try to change details by running below functions
            Badge = int(flexi_badge_input.get())
            employeeID = employee_id_input.get()
            Location = location_input.get()
            etarmis = etarmis_login_input.get()
            change_Attr("EmployeeID", employeeID)
            change_Attr("physicalDeliveryOfficeName", Location)
            change_Attr("badgeID", Badge)
            change_Attr("etarmisNumber", etarmis)
        except:
            MessageBox(None,
                       "There appears to have been an error.\n"
                       "There has been an error sending these details to our "
                       "Active Directory. Please contact your IT Department.",
                       "Error", 0)  # If error, then throw message and close app.

            app.destroy()
            sys.exit()  # exit

        gen_CSV()
        MessageBox(None, "Your details have been successfully updated.", "Success!", 0)  # Details changed - advise.
        app.destroy()
        xmlParse.completed()  # Set counter to 0
        xmlParse.updateInfo()
        sys.exit()  # Exit


def callhelp():
    """ BUILDING HELP """
    helper.overrideredirect(False)
    Q1 = Label(helper, text="Q1: What if I don't know my Employee number?", font=appFontNORM)
    A1 = Label(helper, text="Your employee number can be found at the top of your payslip.\nIt is listed as "
                            "'Employee Reference'")
    Q2 = Label(helper, text="Q2: What's my Etarmis ID?", font=appFontNORM)
    A2 = Label(helper, text="This is optional, if you use the flexi system the number is located on the back of your "
                            "ID card.")
    Q3 = Label(helper, text="Q3: Where can I find my badge ID?", font=appFontNORM)
    A3 = Label(helper, text="This can located on the front of your Council ID badge, just below your name.")
    Q4 = Label(helper, text="Q4: Why are you requesting my personal\nphone number and e-mail?", font=appFontNORM)
    A4 = Label(helper,
               text='We want to make sure we have up-to-date contact information for all employees to allow us\nto '
                    'contact you in an emergency, or get essential information to you quickly, e.g. if we have \nto '
                    'close your work location due to flooding.')
    Q5 = Label(helper, text="Q5: Where is this information stored and for how long?", font=appFontNORM)
    A5 = Label(helper, text="Your personal mobile number and e-mail address will be stored securely in Oracle,\n"
                            "the council’s HR management system. We will keep your personal telephone number \nand "
                            "e-mail address on our record for a period of time for auditing purposes.")

    # Display above text in grid format
    Q1.grid(column=0, row=1, sticky=W, pady=2)
    A1.grid(column=0, row=2, sticky=W, padx=2)

    Q2.grid(column=0, row=3, sticky=W, pady=2)
    A2.grid(column=0, row=4, sticky=W, padx=2)

    Q3.grid(column=0, row=5, sticky=W, pady=2)
    A3.grid(column=0, row=6, sticky=W, padx=2)

    Q4.grid(column=0, row=7, sticky=W, pady=2)
    A4.grid(column=0, row=8, sticky=W, padx=2)

    Q5.grid(column=0, row=9, sticky=W, pady=2)
    A5.grid(column=0, row=10, sticky=W, padx=2)

    helper.resizable(0, 0)  # Lock help box size
    helper.geometry('510x320')  # Window size
    helper.title('Help - Q&A')  # Window title
    helper.mainloop()  # Help main loop

photo = PhotoImage(file=r'path_to_company_logo')  # Assigning .png to photo label
photo_label = Label(app, image=photo, background="white")  # Applying photoimage to label
photo_label.grid(column=1, row=0)  # Applying label to grid

# Labels
opt_label = Label(app, text="       Fields marked \n       * are required.", background="white", font=appFontBOLD)
name_label = Label(app, text="User Login: " + str(user.get_attribute("SAMAccountName"))[2:-2], background="white",
                   font=appFontNORM)
personal_email_label = Label(app, text="Personal E-Mail: ", background="white", font=appFontNORM)
personal_phone_label = Label(app, text="Personal Phone*: ", background="white", font=appFontNORM)
employee_id_label = Label(app, text="Employee Number*: ", background="white", font=appFontNORM)
etarmis_login_label = Label(app, text="Etarmis ID: ", background="white", font=appFontNORM)
flexi_badge_label = Label(app, text="ID Badge Number*: ", background="white", font=appFontNORM)
location_label = Label(app, text="Primary Location*: ", background="white", font=appFontNORM)
info_label1 = Label(app, text="Problems? Contact", background="white", font=appFontBOLD)
info_label2 = Label(app, text=" the ICT Service Desk at: IT SERVICE DESK NUMBER.", background="white", font=appFontBOLD)

# Text Boxes
personal_email_input = Entry(app, width=38)
employee_id_input = Entry(app, width=38)
personal_phone_input = Entry(app, width=38)
etarmis_login_input = Entry(app, width=38)
flexi_badge_input = Entry(app, width=38)

# Combo box
location_input = Combobox(app, width=35, state="readonly")
location = [###ARRAY OF LOCATIONS###]
location_input['values'] = location

# Buttons
submit_button = Button(app, text="Confirm", command=callback)  # assign callback function to submit button
help_button = Button(app, text="Help", command=callhelp)  # assign callhelp function to help button

# GUI items
app.resizable(0, 0)
app.geometry('380x360')  # Window size
app.title('Employee Information Update')  # Window title
app.configure(background="white")
opt_label.grid(column=0, row=1, sticky=E, pady=5)
name_label.grid(column=1, row=1, sticky=N, pady=5)
employee_id_label.grid(column=0, row=2, sticky=E, pady=5, padx=5)
employee_id_input.grid(column=1, row=2, sticky=W)
etarmis_login_label.grid(column=0, row=3, sticky=E, pady=5, padx=5)
etarmis_login_input.grid(column=1, row=3, sticky=W)
flexi_badge_label.grid(column=0, row=4, sticky=E, pady=5, padx=5)
flexi_badge_input.grid(column=1, row=4, sticky=W)
personal_email_label.grid(column=0, row=5, sticky=E, pady=5, padx=5)
personal_email_input.grid(column=1, row=5, sticky=W)
personal_phone_label.grid(column=0, row=6, sticky=E, pady=5, padx=5)
personal_phone_input.grid(column=1, row=6, sticky=W)
location_label.grid(column=0, row=7, sticky=E)
location_input.grid(column=1, row=7, sticky=W, pady=5)
help_button.grid(column=1, row=8, sticky=E, padx=18)
submit_button.grid(column=1, row=8, sticky=W)
info_label1.grid(column=0, row=9, sticky=E)
info_label2.grid(column=1, row=9, sticky=W)

# Arranging Tab Index Order
employee_id_input.lift()
etarmis_login_input.lift()
flexi_badge_input.lift()
personal_email_input.lift()
personal_phone_input.lift()

# xmlParse
xmlParse.initial()  # start initial XML parse
flag = xmlParse.checkFlag()  # assign return value of checkFlag to flag variable
if not xmlParse.checkDate():  # if date < 180 days
    if flag:  # if flag == true
        sys.exit()  # Close app
else:  # If either conditions fail, continue
    pass


def incrementCount():  # Adds count to counter and closes app.
    xmlParse.closedApp()
    sys.exit()


def permission():
    # Explain app to user, if 'No' then close app. If 'yes' then render app.
    answer = MessageBox(None,
                        'Please take the time to update your details with the following form. Any fields that are '
                        'marked with an Asterisk(*) are required, all other fields are optional.',
                        'Employee Information Update', adConnect.win32con.MB_YESNO)
    if answer == 7:  # Answer == no
        if xmlParse.checkCount():  # Check if > 9 count
            pass  # If > 9 pass
        else:  # Else
            incrementCount()  # Count + 1
    else:
        gain_defaults()  # Answer == yes - gain_defaults()


# This logic adds the functionality of ONLY close app if date < 180 and flag is true
if not xmlParse.checkDate():  # If < 180 days
    if flag:  # If flag == true
        sys.exit()  # Exit
    if xmlParse.checkCount():  # If count > 9
        app.overrideredirect(True)  # Remove ability to close
        app.wm_attributes("-topmost", 1)  # Make window always on top.

if xmlParse.checkCount():  # If count > 9
    app.overrideredirect(True)  # Remove ability to close
    app.wm_attributes("-topmost", 1)  # Make window always on top

permission()  # Start app with messagebox
app.mainloop()  # Main app loop

import requests
from bs4 import BeautifulSoup
from time import sleep
import json
from datetime import datetime
import win32com.client as win32
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

sender_email_address = ""
sender_email_password = ""
recipient_email_address = ""
desired_location = ""




def configure_stored_credentials():
    user_input = input("Would you like to restore previous settings, or setup new credentials?\n1.Use stored credentials.\n2.Setup new credentials.\n3.Exit\n")
    while (user_input not in ['1', '2', '3']) or (len(user_input)!=1):
        user_input = input("Please try again. Please only put in the number of your response, and press enter.\n")
    if user_input == '3': #Close application
        print('Closing application.')
        return False
    elif user_input == '1': #Import Previous Settings
        global sender_email_address, sender_email_password, recipient_email_address, desired_location
        load_dotenv()
        sender_email_address = os.getenv('SENDER_EMAIL_ADDRESS')
        sender_email_password = os.getenv('SENDER_EMAIL_PASSWORD')
        recipient_email_address = os.getenv('RECIPIENT_EMAIL_ADDRESS')
        desired_location = os.getenv('DESIRED_LOCATION')
        print (f'Sending email: {sender_email_address}\nSending password: {sender_email_password}\nRecipient Email/Phone: {recipient_email_address}')
        return True
    elif user_input == '2': #Setup new credentials
        #setup sending email
        sender_email_address = input("Please input sending email address.\n")
        while ('@' not in sender_email_address) or (sender_email_address.find('.', sender_email_address.find('@')) == -1):
            sender_email_address = input('Invalid email, please try again.\n')
        sender_email_password = input("Please input sending email password.\n")
        #setup sending email password
        while sender_email_password == "":
            sender_email_password = input ('Please input sending email password. Password field cannot be left blank')
        #determine if sending to email or cell phone
        recipient_email_address = input("Will this be sent to a cell phone or email?\n1.Email\n2.Cell Phone\n")
        while (recipient_email_address not in ['1', '2']) or (len(recipient_email_address)!=1): 
            recipient_email_address = input("Please try again. Please only put in the number of your response, and press enter.\n")
        #if sending to email
        if recipient_email_address == '1':
            recipient_email_address = input("Please input recipient email address.\n")
            while  ('@' not in recipient_email_address) or (recipient_email_address.find('.', recipient_email_address.find('@')) == -1):
                recipient_email_address = input('Invalid email, please try again.\n')
        #if sending to cell phone
        elif recipient_email_address == '2':
            #get phone number
            recipient_email_address = input("Please input 10 digit recipient cell phone number including area code.\n")
            chars_to_remove = "()-. "
            for item in chars_to_remove:
                recipient_email_address = recipient_email_address.replace(item,'')
            while (recipient_email_address.isdigit()==False) or (len(recipient_email_address)!=10):
                recipient_email_address = input("Invalid phone number, please try again.")
                chars_to_remove = "()-. "
                for item in chars_to_remove:
                    recipient_email_address = recipient_email_address.replace(item,'')
            #get cell phone carrier
            #convert cell number to email
        else:
            print ('Unexpected error, closing out application.')
        print (f'Sending email: {sender_email_address}\nSending password: {sender_email_password}\nRecipient Email/Phone: {recipient_email_address}')
        #verify range 1 or 2, then we need to convert phone numbers
        return True
    else:
        print ('Unexpected error, closing out application.')
        return False

    # if user_input not in range(3):
    # if user_input == 1:
    # elif user_input ==2:
    # elif user_input ==3:
    # else:

def check_availability():
    my_response =requests.post("https://telegov.egov.com/lc_ui/CustomerCreateAppointments/SelectType")
    my_response =requests.post("https://telegov.egov.com/lc_ui/AppointmentWizard/61")
    my_response = my_response.text
    soup = BeautifulSoup(my_response, 'lxml')
    try:
        return soup.find('div', class_= "badge badge-pill badge-danger text-wrap mb-2").text
    except:
        return ("Availability")

def list_openings():
    my_response =requests.post("https://telegov.egov.com/lc_ui/AppointmentWizard/61")
    my_response = my_response.text
    soup = BeautifulSoup(my_response, 'lxml')
    data = soup.find_all("script", type="text/javascript")[3].get_text()
    my_json = (data[data.find("["):data.find("]")+1])
    my_json_var = json.loads(my_json)
    for item in my_json_var:
        if item['IsFullyBooked'] == False:
            now = datetime.now()
            print (f"{now}: {item['City']} has an opening")
            if item['City']==desired_location:
                send_email_notification(recipient_email_address)
                return True
            

def send_email_notification(outbound_email):
    now = datetime.now()
    msg = EmailMessage()
    msg['From'] = sender_email_address
    msg['To'] = outbound_email
    msg['Subject'] = 'Appointment Opening'
    msg.set_content(f"Current appointment availablility now at {desired_location} for the unemployment office.\nCurrent time of opening is: {now}\n Sign up now at the following link: https://telegov.egov.com/lc_ui/AppointmentWizard/61")
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email_address, sender_email_password)
        server.send_message(msg)
        server.quit()
        print ('Email Sent!')
        sleep(300)
    except:
        print ('Something went wrong when attempting to send email')

def run_script(my_bool):
    while my_bool == True:
        configure_stored_credentials()
        now = datetime.now()
        if check_availability()==('No Availability'):
            print ('--------------------')
            print (f"{now}: no appointments at this time")
            sleep(6)
        else:
            print ('--------------------')
            print (f'{now}: Openings found at below locations: ')
            if list_openings() == True:
                my_bool = False
            sleep (30)
            
#run_script(True)
configure_stored_credentials()
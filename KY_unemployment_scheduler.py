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


load_dotenv()
sender_email_address = os.getenv('SENDER_EMAIL_ADDRESS')
sender_email_password = os.getenv('SENDER_EMAIL_PASSWORD')
recipient_email_address = os.getenv('RECIPIENT_EMAIL_ADDRESS')
desired_location = os.getenv('DESIRED_LOCATION')

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
            
run_script(True)
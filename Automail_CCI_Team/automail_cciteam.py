#Warning : If you want to successfully run this script, please connect with Google Sheet API from this tutorial :  https://www.youtube.com/watch?v=cnPlKLEGR7E
# and create a password for your gmail from this tutorial : https://www.youtube.com/watch?v=Bg9r_yLk7VY&t=776s

#This Script will take data from your Google Sheet cells and send a custom mail with those data
#I use this script to send task for my employee. 
# Basically, I have to send mail to 100 employees with the same email template but with different tasks name, links and deadline

#..........Code..............#

import time
import smtplib #Connecting Gmail
import gspread #Connecting GoogleSheet
from oauth2client.service_account import ServiceAccountCredentials #Connecting Google Sheet
from datetime import datetime



#This Function will take info from Google Sheet and send email to everyone
# Please change this function parameter according to your needs
# I took user_email to say where this script has to send mail and other staffs to create custom mail
# Check the example.jpg file on this repository for better understanding

def send_email(user_name,user_email,task_name,task_link,deadline) :
    current_time = datetime.now().strftime("%d-%m-%Y") #Not Necessary
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.ehlo()

    server.login('iamakifislam@gmail.com','YourPassword') #Please put your password
    body = f'''

Dear {user_name},

You are assigned to make a Content on a story named "{task_name}." Please read the following document for instruction for details. 

Instructions about making Google Slide : 
(Include a Google Doc Link for Instruction)

Your Task ({task_name}) : 
{task_link}

Deadline : {deadline}


Thanks & Regards
Akif Islam
Young Executive
Tinkers Technologies Ltd

'''
    subject = f"Content Developer's Next Task [ Assigned on : {current_time}]"

    message = f"Subject: {subject}\n\n{body}"
    # print(message);





    server.sendmail(
        'iamakifislam@gmail.com',
        user_email,
        message
    )


    server.quit()
    


def notify_all_cci_members() :
    
    #Connecting Google Sheet API
    
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    
    
    #Checkout the "GoogleAPI_Key"json file. It contains a key to connect API
    
    creds = ServiceAccountCredentials.from_json_keyfile_name("automail_cciteam_googleapi.json", scope);
    
    subscribers = gspread.authorize(creds)
    
    #Please Add Sheet Name 
    sheet = subscribers.open("Content Team Task Asssignments ").sheet1
    
    #Now get info from your required column and save it on a variable
    
    cci_name = sheet.col_values(2)
    cci_email = sheet.col_values(3)
    task_name = sheet.col_values(4)
    cci_task_link = sheet.col_values(5)
    deadline = sheet.col_values(8)
    sub_count = len(cci_task_link)
    
    
    # As it is 0 indexed, I have to subtract 1 row number
    
    start_row = int(input("Start Row of Google Sheet : "))
    end_row = int(input("End Row of Google Sheet : "))

    #As it is a 0 - indexed caluclation :
    start_row = start_row - 1
    end_row = end_row - 1

    while start_row <= end_row:
        #If I don't want to send an email to someone, I just remove the '@' sign from his/her's email. Then script will ignore him/her.
        
        if '@' in cci_email[start_row]:
        
            print(f"Sent To :{cci_name[start_row]} |  Email : {cci_email[start_row]}") #It will print receiver information
        
            send_email(cci_name[start_row],cci_email[start_row],task_name[start_row],cci_task_link[start_row],deadline[start_row]) #This function will send mail
        
            start_row += 1
            
            time.sleep(.5) #I gave 0.5 second break otherwise Gmail will consider you a robot and block you



notify_all_cci_members()
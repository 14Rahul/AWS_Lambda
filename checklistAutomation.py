import boto3
import csv
from datetime import datetime, timedelta
import smtplib
import os
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from openpyxl import load_workbook

s3 = boto3.resource('s3')
bucket = s3.Bucket("sample-lambda-script")



key = "Book1.xlsx"

def lambda_handler(event, context):
    local_file_name = '/tmp/Book1.xlsx' #
    s3.Bucket("sample-lambda-script").download_file(key,local_file_name)
    instanceid = []
    i = 0
    flag = 0
    
    filename1= "/tmp/Magma-Prod"+datetime.now().strftime("%d-%m-%Y_%H-%M-%S")+".xlsx"
    
    client = boto3.client('ec2')
    
    Myec2=client.describe_instances()
    for pythonins in Myec2['Reservations']:
         for printout in pythonins['Instances']:
            #  for a in printout['Tags']:
            #      if (a['Key'] == 'Project') and (a['Value'] == 'TB-PROD'):
                     instanceid.append([])
                     nested = instanceid[i]
                     for d in printout['Tags']:
                        if d['Key'] == 'Name':
                             nested.append(d['Value'])
                     nested.append(printout['InstanceId'])
                     nested.append(printout['InstanceType'])
                     nested.append(printout['State']['Name'])
                     i+=1
    
    
    
    response1 = client.describe_instance_status()
    
    for h in range(i):
        for status in response1['InstanceStatuses']:
            if status['InstanceId'] == instanceid[h][1]:
                nested3 = instanceid[h]
                if (status["InstanceStatus"]["Status"] == "ok") and (status["SystemStatus"]["Status"]== "ok"):
                    nested3.append("Healthy")
                else:
                    nested3.append("Unhealthy")
    
    
    
    
    dbinstance = []
    d = 0
    
    client = boto3.client('rds')
    response2 = client.describe_db_instances()
    for db_instance in response2['DBInstances']:
        if (db_instance['DBInstanceIdentifier'] == 'dev-database-instance-1') or (db_instance['DBInstanceIdentifier'] == 'uatdatabase-instance-1'):
            db_instance_name = db_instance['DBInstanceIdentifier']
            db_type = db_instance['DBInstanceClass']
            db_storage = db_instance['AllocatedStorage']
            db_engine =  db_instance['Engine']
            db_status = db_instance['DBInstanceStatus']
            db_endpoint = db_instance['Endpoint']['Address']
            dbinstance.append([db_instance_name, db_instance_name, db_type, db_endpoint, db_status])
            d+=1
    
    
    for i in range(0,len(instanceid)):
        nested4 = instanceid[i]
        if (instanceid[i][3] == "stopped" or instanceid[i][3] == "terminated"):
            nested4.append("NA")
    
    print(instanceid)
    
    for i in range(0,len(instanceid)):
        if(instanceid[i][3] == "Unhealthy"):
            flag=1
           
    if flag == 0:
        workbook = load_workbook(filename="/tmp/Book1.xlsx")
     
    
        sheet = workbook.active
        
        #modify the desired cell
        sheet["A1"] = "Checklist-Time: "+datetime.now().strftime("%H:%M")
        sheet["A2"] = datetime.now().strftime("%d-%m-%Y")
    
        
        #save the file
        workbook.save(filename1)
    
        body = "Dear Nirnoy\n \nPlease find attached Magma Channel Portal Complete Infra Report with CPU utilization Report. SSH is working.\n \nRegards \nRahul Gupta"
        sender_email = "rahulgupta@nseit.com" 
        receiver_email = "rahulgupta@nseit.com"
        password = ""
    
        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = "M Channel Portal Complete Infra Report"
        message["Cc"] = ""
        #message["Bcc"] = receiver_email  # Recommended for mass emails
    
        # Add body to email
        message.attach(MIMEText(body, "plain"))
    
        filename = [ filename1 ]  # In same directory as script
    
        dir_path = "."
        for f in filename:  # add files to the message
                file_path = os.path.join(dir_path, f)
                attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
                attachment.add_header('Content-Disposition','attachment', filename=f)
                message.attach(attachment)
    
        text = message.as_string()
    
        # Log in to server using secure context and send email
        context = ssl._create_unverified_context()
        with smtplib.SMTP_SSL("nseit.icewarpcloud.in", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email.split(",") + message["Cc"].split(","), text)
        
        folder_path = ('.')
        os.remove(filename1)
    else:
        body = "Instance unhealthy Magma"
        sender_email = "rahulgupta@nseit.com" 
        receiver_email = "rahulgupta@nseit.com"
        password = ""
    
        # Create a multipart message and set headers
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = "Unhealthy Error"
        message["Cc"] = ""
        #message["Bcc"] = receiver_email  # Recommended for mass emails
    
        # Add body to email
        message.attach(MIMEText(body, "plain"))
    
        text = message.as_string()
    
        # Log in to server using secure context and send email
        context = ssl._create_unverified_context()
        with smtplib.SMTP_SSL("nseit.icewarpcloud.in", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email.split(",") + message["Cc"].split(","), text)
        print("Khatam")


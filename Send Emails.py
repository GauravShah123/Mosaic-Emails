# Import modules
import smtplib, ssl
## email.mime subclasses
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
## The pandas library is only for generating the current date, which is not necessary for sending emails
import pandas as pd

import csv
import os

def list_files():
    files = os.listdir(".")  # Get all files in the current directory
    print("\nFiles in the current folder:")
    for file in files:
        print(f"- {file}")
    print()  # Add a new line for better readability

list_files()

# Set up the email addresses and password. Please replace below with your email address and password
email_from = 'gbdasoc@gmail.com'
password = 'gepj ifjm uhki hlje'

# Generate today's date to be included in the email Subject
date_str = pd.Timestamp.today().strftime('%Y-%m-%d')


# Get CSV filename from user
filename = input("Enter the CSV filename (including .csv extension): ")
htmlfile = input("Enter the HTML template filename (including .html extension): ")

# Initialize empty lists
names = []
emails = []

# Read the CSV file
try:
    with open(filename, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)  # Using DictReader to handle column names
        for row in reader:
            names.append(row["Name"])   # Extracting names
            emails.append(row["Email"]) # Extracting emails

    # Print results (optional)
    print("Names:", names)
    print("Emails:", emails)

except FileNotFoundError:
    print("Error: File not found. Please make sure the file is in the same folder.")
except KeyError:
    print("Error: Ensure your CSV has 'Name' and 'Email' columns.")

# Get the HTML content
def get_email_content(name, filename):
    try:
        with open(filename, "r", encoding="utf-8") as file:
            html_content = file.read()  # Read the HTML file
            return html_content.replace("{{NAME}}", name)  # Replace placeholder
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
        return None

subject = input("Subject: ")

# Define a function to attach files as MIMEApplication to the email
def attach_file_to_email(email_message, filename):
    # Open the attachment file for reading in binary mode, and make it a MIMEApplication class
    with open(filename, "rb") as f:
        file_attachment = MIMEApplication(f.read())
    # Add header/name to the attachments    
    file_attachment.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    # Attach the file to the message
    email_message.attach(file_attachment)    

attachments = input("List attachments separated by commas (or leave empty): ")

def get_file_attachments(filenames: str):
    """Convert a comma-separated string of filenames into a list."""
    if not filenames.strip():  # Check if the input is empty or just spaces
        return []
    return [filename.strip() for filename in filenames.split(',')]

input("Good to send?")

for i in range(len(names)):
    email_to = emails[i]

    # Create a MIMEMultipart class, and set up the From, To, Subject fields
    email_message = MIMEMultipart()
    email_message['From'] = email_from
    email_message['To'] = email_to
    email_message['Subject'] = subject

    # Attach the HTML content
    email_message.attach(MIMEText(get_email_content(names[i].split()[0], htmlfile), "html"))

    # Get the list of attachments
    file_list = get_file_attachments(attachments)

    # Attach files if any exist
    if file_list:
        for file in file_list:
            attach_file_to_email(email_message, file)
    else:
        print("No attachments provided.")

    # Convert it as a string
    email_string = email_message.as_string()

    # Connect to the Gmail SMTP server and Send Email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(email_from, password)
        server.sendmail(email_from, email_to, email_string)

    print(f"Email sent to {names[i]}")
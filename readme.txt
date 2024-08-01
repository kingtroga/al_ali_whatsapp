README.txt


Overview
This script sends a WhatsApp message with an image to a list of phone numbers, specifically targeting parents. It filters out group admins from the list and handles errors by logging unsuccessful attempts.

Prerequisites
Python 3.x
pywhatkit library
win32clipboard library
pandas library
Excel files: AL-ALI.xlsx, messaged_numbers.xlsx, and error_numbers.xlsx
Installation
Install the required Python libraries:


pip install pywhatkit pywin32 pandas openpyxl


Ensure you have the necessary Excel files in the same directory as the script:

AL-ALI.xlsx: Contains the phone numbers and admin status.
messaged_numbers.xlsx: Tracks the numbers that have already been messaged.
error_numbers.xlsx: Tracks the numbers where messaging failed.
Usage
Prepare the environment:

Make sure AL-ALI.xlsx exists and contains the columns PHONE NUMBERS and IS ADMIN.
Ensure you have the image file images/2.jpg in the specified path.
The script sends the following message:
text

Al Ali International School Summer Camp.
Starting August 5th, 2024.
Open to the general public.


Run the script:

Execute the script in a Python environment. It will:
Filter out admins from AL-ALI.xlsx and save the filtered list in filtered_file.xlsx.
Send the message to up to 20 phone numbers at a time.
Log successful and failed attempts in messaged_numbers.xlsx and error_numbers.xlsx respectively.
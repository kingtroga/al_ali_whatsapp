import pywhatkit
import win32clipboard
import pandas as pd
import os

# Load messaged numbers
if os.path.exists('messaged_numbers.xlsx'):
    df = pd.read_excel('messaged_numbers.xlsx')
    if df.empty:
        messaged_numbers = []
    else:
        messaged_numbers = df['MESSAGED NUMBERS'].to_list()
else:
    messaged_numbers = []

# Load error numbers
if os.path.exists('error_numbers.xlsx'):
    df = pd.read_excel('error_numbers.xlsx')
    if df.empty:
        error_numbers = []
    else:
        error_numbers = df['ERROR NUMBERS'].to_list()
else:
    error_numbers = []

CAPTION = """Al Ali International School Summer Camp.
Starting August 5th, 2024.
Open to the general public."""

# Remove the group admins from the phone numbers
if not os.path.exists('filtered_file.xlsx'):
    try:
        df = pd.read_excel('AL-ALI.xlsx')
        filtered_df = df[df['IS ADMIN'] != True]
        filtered_df.to_excel('filtered_file.xlsx', index=False)
    except Exception as e:
        print(f"Error while processing AL-ALI.xlsx: {e}")
        raise

# Load the filtered phone numbers
try:
    df = pd.read_excel('filtered_file.xlsx')
except Exception as e:
    print(f"Error while reading filtered_file.xlsx: {e}")
    raise

# Send the summer school message to parents only
j = 1
phone_numbers = df['PHONE NUMBERS'].to_list()

#TBH, pywhatkit can only send messages to 20 people at a time before
# it crashes a 16gb ram laptop because it keeps opening new chrome tabs
for phone_number in phone_numbers[0:20]:
    try:
        if phone_number not in messaged_numbers:
            pywhatkit.sendwhats_image(f"+{phone_number}", "images/2.jpg", CAPTION, wait_time=35)
            print(f"Messaged Phone number {j}/{len(phone_numbers)}")
            messaged_numbers.append(phone_number)
            j += 1
    except Exception as e:
        print(f"Error: {e} occurred with +{phone_number}")
        error_numbers.append(phone_number)
        continue

# Save error numbers
try:
    df_error = pd.DataFrame({'ERROR NUMBERS': error_numbers})
    df_error.to_excel("error_numbers.xlsx", index=False)
except Exception as e:
    print(f"Error while saving error_numbers.xlsx: {e}")
    raise

# Save messaged numbers
try:
    df_messaged = pd.DataFrame({'MESSAGED NUMBERS': messaged_numbers})
    df_messaged.to_excel("messaged_numbers.xlsx", index=False)
except Exception as e:
    print(f"Error while saving messaged_numbers.xlsx: {e}")
    raise

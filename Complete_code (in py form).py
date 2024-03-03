from User_Credentials import *
from Contact_List import *
import pyttsx3 as tts
import speech_recognition as sr
import ssl as sl
import smtplib as sm
import imaplib as ilb
from email import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import decode_header
import re
import unidecode


# Used to text to speech

def Text_to_speech(content):
    engine = tts.init()
    # engine.setProperty('rate', 460) -- Use this to adjust the speed of the speech
    engine.say(content)
    engine.runAndWait()
    return

# Used to convert speech to text

def Speech_to_text():
    recogniser = sr.Recognizer()
    while(1):
        with sr.Microphone() as access_to_system_mic:
            recogniser.energy_threshold = 385
            recogniser.adjust_for_ambient_noise(access_to_system_mic,duration=1)
            try:
                audio = recogniser.listen(access_to_system_mic,timeout=15,phrase_time_limit=10)
                text = recogniser.recognize_google(audio)
                return(text)
            except sr.UnknownValueError:
                print('Apologies!!! could not able to hear what you said. Please speak again.')
                Text_to_speech('Apologies!!! could not able to hear what you said.Please speak again.')
            except sr.WaitTimeoutError:
                print('Apologies!!! could not able to hear what you said. Please speak again.')
                Text_to_speech('Apologies!!! could not able to hear what you said.Please speak again.')
            except sr.RequestError as e:
                print("Please make sure that you are connected to the Internet and try again later")
                Text_to_speech("Please make sure that you are connected to the Internet and try again later")

# Actual code for sending emails

def Compose_email(recipients_Email_Id,subject,body):
    print('Sending the Email, please wait.')
    Text_to_speech('Sending the Email, please wait.')
    message = MIMEMultipart()
    message['From'] = sender_Email_Id
    if(len(recipients_Email_Id) > 1):
        message['To'] = ' ,'.join(recipients_Email_Id)
    else:
        message['To'] = ' '.join(recipients_Email_Id)
    message['Subject'] = subject
    message.attach(MIMEText(body,'plain'))
    context = sl.create_default_context()
    try:
        with sm.SMTP_SSL('smtp.gmail.com',465,context = context) as server:
            server.login(sender_Email_Id,sender_AppKey)
            print('success')
            server.sendmail(sender_Email_Id,recipients_Email_Id,message.as_string())
            print('Email Sent Successfully!!!')
            Text_to_speech('Email Sent Successfully')
            print("Sir, I'm done with this task")
            Text_to_speech("Sir, I'm done with this task")
            return
    except:
        print('Error encountered while sending the email!!!\nPlease check the given recipients email id and try again.')
        Text_to_speech('Error encountered while sending the email. Please check the given recipients email id and try again.')
        return

# Auxillary Function for sending_email()

def contact_list():
    
    recipient_emails = []
    print('Sir, once please listen to the contact list')
    Text_to_speech('Sir, once please listen to the contact list')
    for i in contacts:
        print(f'{i}:{contacts[i]}')
        Text_to_speech(f'{i}:{contacts[i]}')
        
    while(1):
        print('Sir, please speak out the recipients you want to include (e.g., "recipient 1 and recipient 3")')
        Text_to_speech('Sir, please speak out the recipients you want to include')
        recipients_input = Speech_to_text().lower().split(' and ')
        for recipient_input in recipients_input:
            for j in contact_lst:
                if any(alias.lower() in recipient_input for alias in [j] + contacts[j].split('@')[0].split('.')):
                    recipient_emails.append(contacts[j])
                    break
                
        if(len(recipient_emails) >= 1):
            return recipient_emails
        else:
            print('Sir, I request you to please repeat the process again as there is some glitch while processing.')
            Text_to_speech('Sir, I request you to please repeat the process again as there is some glitch while processing.')

# Auxillary Function for Compose_email()

def sending_email():
    recipients = contact_list()
    print('Please speak the subject of the Email:')
    Text_to_speech('Please speak the subject of the Email')
    subject = Speech_to_text()
    print(subject)
    print('Please speak the body of the Email')
    Text_to_speech('Please speak the body of the Email')
    body = Speech_to_text()
    print(body)
    Compose_email(recipients,subject,body)

def sanitize_gmail_subject(subject):
    """
    As the subject of an email may contain these special symbols: '/', '\\', '*', ':', '?', '<', '>', '|', 
    it is important to avoid these characters to save the file in the system directory, so this function is 
    used to filter out these special character along eith other characters.
    """   
    if isinstance(subject, bytes):
        subject = subject.decode('utf-8', errors='ignore')

    sanitized_subject = unidecode.unidecode(subject)
    
    sanitized_subject = ''.join(e for e in sanitized_subject if (e.isalnum() or e.isspace()))
    
    forbidden_characters = ['/', '\\', '*', ':', '?', '<', '>', '|']
    
    sanitized_subject = ''.join(char for char in sanitized_subject if char not in forbidden_characters)
    
    sanitized_subject = sanitized_subject.replace('\r', '').replace('\n', '')

    return sanitized_subject

def filter_links(text):
    """
    Filter out links from the given text.
    Returns a tuple containing the filtered text and a flag indicating if links were present.
    """
    url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    filtered_text = re.sub(url_pattern, '', text)

    links_present = filtered_text != text

    return filtered_text, links_present

# Used to fetch and read the first five emails

def inspect_email():

    mail_obj = ilb.IMAP4_SSL("imap.gmail.com")
    mail_obj.login(sender_Email_Id, sender_AppKey)

    print('Sir, please listen to the number which is associated with the Gmail labels')
    Text_to_speech('Sir, please listen to the number which is associated with the Gmail labels')
    print('Number 1: Inbox')
    Text_to_speech('Number 1: Inbox')
    print('Number 2: Starred')
    Text_to_speech('Number 2: Starred')
    print('Number 3: Important')
    Text_to_speech('Number 3: Important')
    print('Number 4: Sent')
    Text_to_speech('Number 4: Sent')
    print('Number 5: Drafts')
    Text_to_speech('Number 5: Drafts')
    print('Number 6: Spam')
    Text_to_speech('Number 6: Spam')
    print('Number 7: Bin')
    Text_to_speech('Number 7: Bin')
    print('Sir, please select the number which is associated with the label which you want to access:')
    Text_to_speech('Sir, please select the number which is associated with the label which you want to access:')
    user_choice = Speech_to_text()
    print(user_choice)
    if('number 1' in user_choice or 'number one' in user_choice or 'number won' in user_choice or 'number van' in user_choice):
        status_of_selection_state, messages = mail_obj.select('Inbox')
    elif('number 2' in user_choice or 'number tu' in user_choice or 'number two' in user_choice or 'number to' in user_choice or 'number too' in user_choice):
        status_of_selection_state, messages = mail_obj.select('[Gmail]/Starred')
    elif('number 3' in user_choice or 'number three' in user_choice):
        status_of_selection_state, messages = mail_obj.select('[Gmail]/Important')
    elif('number 4' in user_choice or 'number four' in user_choice or 'number for' in user_choice):
        status_of_selection_state, messages = mail_obj.select('"[Gmail]/Sent Mail"')
    elif('number 5' in user_choice or 'number five' in user_choice):
        status_of_selection_state, messages = mail_obj.select('[Gmail]/Drafts')
    elif('number 6' in user_choice or 'number six' in user_choice):
        status_of_selection_state, messages = mail_obj.select('[Gmail]/Spam')
    elif('number 7' in user_choice or 'number seven' in user_choice):
        status_of_selection_state, messages = mail_obj.select('[Gmail]/Bin')
    else:
        print('As the user choice is invalid, selecting the inbox label as default\n')
        Text_to_speech('As the user choice is invalid, selecting the inbox label as default')
        status_of_selection_state, messages = mail_obj.select('inbox')
    if status_of_selection_state != 'OK':
        print(f"Failed to select mailbox. Status: {status_of_selection_state}")
        Text_to_speech("Failed to select mailbox.")
        return
    else:
        status_of_searching_mailbox, email_charset = mail_obj.search(None, 'ALL')
        email_byte_data = email_charset[0].split()
        first_five_latest_email_byte_data = email_byte_data[-5:]

        for email_ids in reversed(first_five_latest_email_byte_data):
            status_of_fetch_operation_on_email_ids, msg_data = mail_obj.fetch(email_ids, '(RFC822)')
            raw_byte_format_of_fetched_data = msg_data[0][1]
            email_content_in_string_obj_format = message_from_bytes(raw_byte_format_of_fetched_data)

            subject = decode_header(email_content_in_string_obj_format['Subject'])[0][0]

            if email_content_in_string_obj_format.is_multipart():
                print("Subject: ", sanitize_gmail_subject(subject))
                Text_to_speech(f"Subject: {sanitize_gmail_subject(subject)}")
                print(f"From: {decode_header(email_content_in_string_obj_format['From'])[0][0]}")
                Text_to_speech(f"From: {decode_header(email_content_in_string_obj_format['From'])[0][0]}")
                print(f"Date: {decode_header(email_content_in_string_obj_format['Date'])[0][0]} Greenwich Mean Time Zone")
                Text_to_speech(f"Date: {decode_header(email_content_in_string_obj_format['Date'])[0][0]} Greenwich Mean Time Zone")
                for part in email_content_in_string_obj_format.walk():
                    content_disposition = part.get_content_disposition()

                    if content_disposition is not None and ('attachment' in content_disposition):
                        filename = part.get_filename()
                        if filename:
                            filename = decode_header(filename)[0][0]
                            if isinstance(filename, bytes):
                                filename = filename.decode('utf-8')
                            output_location = os.path.join(default_path_to_save_attachments, sanitize_gmail_subject(subject), filename)
                            os.makedirs(os.path.dirname(output_location), exist_ok=True)
                            with open(output_location, 'wb') as attachment_file:
                                attachment_file.write(part.get_payload(decode=True))
                            print(f'Sir, I encountered an attachment named "{filename}" and this attachment is saved in the provided default location. Kindly pay a visit to the specified location.')
                            Text_to_speech(f'Sir, I encountered an attachment named "{filename}" and this attachment is saved in the provided default location. Kindly pay a visit to the specified location.')
                            print("/" * 150)

                    if part.get_content_type() == 'text/plain' and (content_disposition is None or 'attachment' not in content_disposition):
                        body_text, links_present = filter_links(part.get_payload(decode=True).decode('utf-8'))
                        if links_present:
                            print("Sir, the following email contains links. Please visit them.")
                            Text_to_speech("Sir, the following email contains links. Please visit them.")
                            Text_to_speech("Now reading the body of the email without links")
                        print('Body:\n', body_text)
                        Text_to_speech(f'Body: {body_text}')
                        print("/" * 150)

            else:
                print("Subject: ", sanitize_gmail_subject(subject))
                Text_to_speech(f"Subject: {sanitize_gmail_subject(subject)}")
                print(f"From: {decode_header(email_content_in_string_obj_format['From'])[0][0]}")
                Text_to_speech(f"From: {decode_header(email_content_in_string_obj_format['From'])[0][0]}")
                print(f"Date: {decode_header(email_content_in_string_obj_format['Date'])[0][0]} Greenwich Mean Time Zone")
                Text_to_speech(f"Date: {decode_header(email_content_in_string_obj_format['Date'])[0][0]} Greenwich Mean Time Zone")
                Text_to_speech("Sir, I encountered a HTML integrated email, so I'm saving it in the provided default location. Kindly pay a vist to the specified location.")
                print("HTML INTEGRATED EMAIL")
                file_name = f'{sanitize_gmail_subject(subject)}.html'
                if file_name:
                    output_location = os.path.join(default_path_to_save_html_integrated_emails, file_name)
                    os.makedirs(os.path.dirname(output_location), exist_ok=True)
                    with open(output_location, 'w', encoding='utf-8') as op:
                        op.write(email_content_in_string_obj_format.get_payload(decode=True).decode('utf-8','ignore'))
                    print("/" * 150)

        mail_obj.logout()
        
        print("Sir, I'm done with this task")
        Text_to_speech("Sir, I'm done with this task")
        return

# Driver Code

print('Welcome Sir!')
Text_to_speech('Welcome Sir!')
while(1):
    
    print('Sir, Please listen to the following list of operations: ')
    Text_to_speech('Sir, Please listen to the following list of operations: ')
    
    print('Option 1: Compose an Email')
    Text_to_speech('Option 1: Compose an Email')
    
    print('Option 2: Read first five Emails')
    Text_to_speech('Option 2: Read first five Emails')
    
    print('Option 3: Exit the Program')
    Text_to_speech('Option 3: Exit the Program')
    
    print('Sir, I request you to speak out a valid operation.')
    Text_to_speech('Sir, I request you to speak out a valid operation.')
    user_input = Speech_to_text()
    
    if('option 1' in user_input or 'option one' in user_input or 'option won' in user_input or 'option van' in user_input):
        sending_email()
    elif('option 2' in user_input or 'option to' in user_input or 'option too' in user_input or 'option tu' in user_input or 'option two' in user_input):
        inspect_email()
    elif('option 3' in user_input or 'option three' in user_input or 'free' in user_input):
        print('Exiting, please wait.....')
        print('Thank You Sir, Please visit again.')
        Text_to_speech('Thank You Sir, Please visit again, Have a great day sir.')
        break
    else:
        print('Sir, You selected an invalid operation.')
        Text_to_speech('Sir, You selected an invalid operation. So, ')
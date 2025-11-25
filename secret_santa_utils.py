#########################################################################
#------------------------------- IMPORTS -------------------------------#
#########################################################################

import random
import pandas as pd
import win32com.client
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import keyring
import base64
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import requests
import urllib.parse



#########################################################################
#------------------------------- ALGORITHM -----------------------------#
#########################################################################

def assign_matches(participant_list, seed=None):
    """
    Randomly assign each participant a Secret Santa partner.

    Parameters
    ----------
    participant_list : list of str
        A list of participant names.
    seed : int, optional
        Random seed for reproducibility. If provided, the same assignments
        will be generated each time the function is called with the same seed.

    Returns
    -------
    dict
        A dictionary mapping each participant (gifter) to another participant (giftee).
        No participant is assigned to themselves, and each participant both gives
        and receives exactly one gift.

    Examples
    --------
    >>> assign_matches(["Jesse", "Derek", "Angela", "Will"], seed=42)
    {'Jesse': 'Derek', 'Derek': 'Will', 'Angela': 'Jesse', 'Will': 'Angela'}
    """
    
    if seed is not None:
        random.seed(seed)

    # shuffle participant list
    participants_shuffled = participant_list.copy()
    random.shuffle(participants_shuffled)

    # assign matches
    giftees = participants_shuffled[1:] + participants_shuffled[:1]
    
    # build output dictionary
    assignments = {p: giftees[participants_shuffled.index(p)] for p in participant_list}

    return assignments



#########################################################################
#------------------------------- COMMUNICATION -------------------------#
#########################################################################

def get_initial_token(email_address):
    REDIRECT_URI = "http://localhost:8080"
    SCOPE = "https://mail.google.com/"
    client_id = keyring.get_password("gmail_oauth_id", email_address)
    client_secret = keyring.get_password("gmail_oauth_secret", email_address)
    auth_url = (
        "https://accounts.google.com/o/oauth2/v2/auth?"
        + urllib.parse.urlencode({
            "client_id": client_id,
            "redirect_uri": REDIRECT_URI,
            "response_type": "code",
            "scope": SCOPE,
            "access_type": "offline",   
            "prompt": "consent"    
        })
    )
    
    print("Go to this URL and authorize. It will open to a 'Site can't be reached' page. Copy code in URL between ?code= and &scope=.\n", auth_url)

    code = input("Enter the authorization code: ")

    token_url = "https://oauth2.googleapis.com/token"
    data = {
        "code": code,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code"
    }

    response = requests.post(token_url, data=data)
    tokens = response.json()
    print(tokens)



def email_secret_santa_draw(email_dict, giftee_dict, subject=None, group_name=None, instructions=None, error_folder=None, use_outlook=False, sender_email=None):
    """
    Send Secret Santa assignment emails using Outlook if available,
    otherwise fall back to console output.

    Parameters
    ----------
    email_dict : dict
        Dictionary mapping participant names to their email addresses.
    giftee_dict : dict
        Dictionary mapping participant names to their assigned giftee names.

    Returns
    -------
    None
        Sends or prints each participant's Secret Santa assignment.
    """
    
    if not subject:
        subject = "Secret Santa Draw"
    
    try:
        if not use_outlook:
            raise Exception("Outlook usage disabled.")
        
        # send email to each participant
        print("Sending emails via Outlook...")
        for participant, giftee in giftee_dict.items():
            try:
                recipient_email = email_dict.get(participant)
                if not recipient_email:
                    raise ValueError(f"Email address not found for participant: {participant}.")

                outlook = win32com.client.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = recipient_email
                mail.Subject = subject
                mail.HTMLBody = f"""
                    <body>
                        <p>Hi {participant},</p>
                        <p>The {f"{group_name} " if group_name else ""}Secret Santa draw is ready!</p>
                        <p>You are the Secret Santa for: <strong>{giftee}</strong>!</p>
                        {f"<p>Instructions:<br>{instructions}</p>" if instructions else ""}
                    </body>
                    <p><strong>MERRY CHRISTMAS!</strong></p>
                    """

                mail.Send()
                print(f"Email sent to {participant} at {recipient_email}.")

            except Exception as e:
                print(f"Error sending email to {participant} at {recipient_email}: {e}")
                if error_folder:
                    if not os.path.exists(error_folder):
                        os.makedirs(error_folder)
                    with open(os.path.join(error_folder, f"{participant}_giftee.txt"), "w") as f:
                        f.write(f"{giftee}")

        print("Emails sent via Outlook.")
        
        

    except Exception as e:
        print("Sending emails via SMTP...")

        if not sender_email:
            raise ValueError("SMTP email method requires sender_email and sender_password.")

        client_id = keyring.get_password("gmail_oauth_id", sender_email)
        client_secret = keyring.get_password("gmail_oauth_secret", sender_email)
        refresh_token = keyring.get_password("gmail_oauth_token", sender_email)
        
        # build credentials object
        creds = Credentials(
            None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=["https://mail.google.com/"]
        )

        # refresh to get a valid access token
        creds.refresh(Request())
        access_token = creds.token

        # build XOAUTH2 authentication string
        auth_string = f"user={sender_email}\1auth=Bearer {access_token}\1\1"
        auth_bytes = base64.b64encode(auth_string.encode("utf-8")).decode("utf-8")

        # connect to Gmail SMTP
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp_server:
            smtp_server.starttls()
            smtp_server.docmd("AUTH", "XOAUTH2 " + auth_bytes)
            
            # send email to each participant
            for participant, giftee in giftee_dict.items():
                try:
                    recipient_email = email_dict.get(participant)
                    if not recipient_email:
                        raise ValueError(f"Email address not found for participant: {participant}.")

                    body_html = f"""
                        <html>
                            <body>
                                <p>Hi {participant},</p>
                                <p>The {f"{group_name} " if group_name else ""}Secret Santa draw is ready!</p>
                                <p>You are the Secret Santa for: <strong>{giftee}</strong>!</p>
                                {f"<p>Instructions:<br>{instructions}</p>" if instructions else ""}
                            </body>
                            <p><strong>MERRY CHRISTMAS!</strong></p>
                        </html>
                        """
                    body_text = f"""
                        Hi {participant}, 
                        
                        The {f"{group_name} " if group_name else ""}Secret Santa draw is ready!
                        
                        You are the Secret Santa for: {giftee}
                        
                        {f"Instructions:{instructions}" if instructions else ""}

                        MERRY CHRISTMAS!
                        """
                        
                    msg = MIMEMultipart("alternative")
                    msg.attach(MIMEText(body_text, "plain"))
                    msg.attach(MIMEText(body_html, "html"))
                    msg["From"] = sender_email
                    msg["To"] = recipient_email
                    msg["Subject"] = subject

                    smtp_server.sendmail(sender_email, recipient_email, msg.as_string())
                    print(f"Email sent to {participant} at {recipient_email}.")
                    
                except Exception as e:
                    print(f"Error sending email to {participant} at {recipient_email}: {e}")
                    if error_folder:
                        if not os.path.exists(error_folder):
                            os.makedirs(error_folder)
                        with open(os.path.join(error_folder, f"{participant}_giftee.txt"), "w") as f:
                            f.write(f"{giftee}")
            
            smtp_server.quit() 
        print("Emails sent via SMTP.")
        
        
        
        
#########################################################################
#------------------------------- INPUT VALIDATION ----------------------#
#########################################################################

def validate_participants(participant_list):
    """
    Validate the participant list for duplicates and minimum count.

    Parameters
    ----------
    participant_list : list of str
        A list of participant names.
        
    Returns
    -------    
    None
        Raises ValueError if validation fails.
    """
    # check for duplicate names and raise error
    if len(participant_list) != len(set(participant_list)):
        raise ValueError("Duplicate participant names detected. Please ensure all names are unique.")
    if len(participant_list) == 0:
        raise ValueError("No participants found. Please add at least two participants.")
    if len(participant_list) == 1:
        raise ValueError(f"Only one participant found ({participant_list[0]}). Please add at least two participants.")
 


#########################################################################
#------------------------------- UTILITIES -----------------------------#
#########################################################################

def extract_participant_list(excel_file_path, participant_dict):
    """
    Extract participant names and email addresses from an Excel file.

    Parameters
    ----------
    excel_file_path : str
        Path to the Excel file containing participant data.

    Returns
    -------
    dict
        A dictionary mapping participant names to their email addresses.
    """
    if excel_file_path:
        try:
            # Read the Excel file
            df = pd.read_excel(excel_file_path)
            # Ensure the required columns exist
            if 'Name' not in df.columns or 'Email' not in df.columns:
                raise ValueError("File must contain 'Name' and 'Email' columns.")
            # Populate the PARTICIPANTS dictionary
            participants = dict(zip(df['Name'].str.strip(), df['Email'].str.strip()))
            print(pd.DataFrame.from_dict(participants, orient='index', columns=['Email']))
            return participants

        except Exception as e:
            print(f"Attempt to use Excel file for participant list failed: {e}")
            print("Falling back to hardcoded PARTICIPANTS dictionary.")
    
    else:
        return participant_dict

# %%

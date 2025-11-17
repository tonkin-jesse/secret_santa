#########################################################################
#------------------------------- IMPORTS -------------------------------#
#########################################################################

import random
import math
import pandas as pd
import sys
import win32com.client
import smtplib
from email.mime.text import MIMEText

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

def email_secret_santa_draw(email_dict, giftee_dict, subject=None, use_outlook=True, sender_email=None, sender_password=None):
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
        
        # try to connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI") 

        # send email to each participant
        try:
            for participant, giftee in giftee_dict.items():
                recipient_email = email_dict.get(participant)
                if not recipient_email:
                    continue

                mail = outlook.CreateItem(0) 
                mail.To = recipient_email
                mail.Subject = subject
                mail.Body = f"Hi {participant},\n\nYou are the Secret Santa for: {giftee}!\n\nMERRY CHRISTMAS!"
                mail.Send()
                
                print(f"Email sent to {participant} at {recipient_email}.")
        except Exception as e:
            print(f"Error sending email to {participant} at {recipient_email} via Outlook: {e}")

        print("Emails sent via Outlook.")

    except Exception as e:
        print("Outlook not available, falling back to SMTP...")

        if not sender_email or not sender_password:
            raise ValueError("SMTP fallback requires sender_email and sender_password.")

        # connect to Office365 SMTP
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)

            # send email to each participant
            for participant, giftee in giftee_dict.items():
                recipient_email = email_dict.get(participant)
                if not recipient_email:
                    continue

                subject = subject
                body = f"Hi {participant},\n\nYou are the Secret Santa for: {giftee}!\n\nMERRY CHRISTMAS!"

                msg = MIMEText(body)
                msg["From"] = sender_email
                msg["To"] = recipient_email
                msg["Subject"] = subject

                server.sendmail(sender_email, recipient_email, msg.as_string())

        print("Emails sent via SMTP.")
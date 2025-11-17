#########################################################################
#------------------------------- IMPORTS -------------------------------#
#########################################################################

from secret_santa_config import *
from secret_santa_utils import *
import win32com.client
import os
import keyring

#########################################################################
#------------------------------- MAIN ----------------------------------#
#########################################################################

if __name__ == "__main__":
    
    # check for duplicate names and raise error
    if len(PARTICIPANTS) != len(set(PARTICIPANTS)):
        raise ValueError("Duplicate participant names detected. Please ensure all names are unique.")

    # assign gift matches
    secret_santa_draw_dict = assign_matches(list(PARTICIPANTS.keys()))
    
    # send emails
    email_secret_santa_draw(PARTICIPANTS, secret_santa_draw_dict, use_outlook=False, sender_email="jesse.tonkin1999@gmail.com", sender_password=keyring.get_password("EmailSMTP", "jesse.tonkin1999@gmail.com"))




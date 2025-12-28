#########################################################################
#------------------------------- IMPORTS -------------------------------#
#########################################################################

import os
import keyring
USER = os.getlogin()



#########################################################################
#------------------------------- FILES ---------------------------------#
#########################################################################

# provide a link to a excel file with participant names and email addresses or input into PARTICIPANTS below
PARTICIPANT_EXCEL_FILE = None
PARTICIPANT_TXT_FILE = fr"J:\Documents\Code\secret_santa\Participants.txt"
ERROR_OUTPUT_FOLDER = None



#########################################################################
#------------------------------- INPUTS --------------------------------#
#########################################################################

PARTICIPANTS = {
    #"Name": "Email"
}

# Email settings
EMAIL_SUBJECT = "SECRET SANTA DRAW 1 (HOMEMADE/REPURPOSED PRESENT)"
GROUP_NAME = "Tonkins and Special Friends"
INSTRUCTIONS = "&emsp;Something homemade or that you found around the house<br>&emsp;Price range: <=$20<br>&emsp;Exchange date: 25/12/2025" #<br>&emsp;Exchange location: 
USE_OUTLOOK = False
SENDER_EMAIL = "jesse.tonkin1999@gmail.com"





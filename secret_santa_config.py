#########################################################################
#------------------------------- IMPORTS -------------------------------#
#########################################################################

import os
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
    #"Name": ["Email", ["AntiSecretSanta1", "AntiSecretSanta2"]],
}

# Email contents
EMAIL_SUBJECT = "SECRET SANTA"
GROUP_NAME = "Tonkins and Special Friends"
INSTRUCTIONS = "&emsp;Something homemade or that you found around the house<br>&emsp;Price range: <=$20<br>&emsp;Exchange date: 25/12/2025" #<br>&emsp;Exchange location: 

# Email settings
SENDER_EMAIL = os.environ["EmailAddress"]
USE_OUTLOOK = False






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
ERROR_OUTPUT_FOLDER = None



#########################################################################
#------------------------------- INPUTS --------------------------------#
#########################################################################

PARTICIPANTS = {
    "Jesse Tonkin": "jesse.tonkin1999@gmail.com",
    "Jesse Tonkin2": "jesse.b.tonkin@gmail.com",
}

# Email settings
EMAIL_SUBJECT = "SECRET SANTA"
GROUP_NAME = "Team"
INSTRUCTIONS = "&emsp;Price range: $15-20<br>&emsp;Exchange date: <br>&emsp;Exchange location: "
USE_OUTLOOK = False
SENDER_EMAIL = "jesse.tonkin1999@gmail.com"






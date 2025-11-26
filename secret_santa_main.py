#%%#########################################################################
#------------------------------- IMPORTS -------------------------------#
#########################################################################

from secret_santa_config import *
from secret_santa_utils import *

#%%

#########################################################################
#------------------------------- MAIN ----------------------------------#
#########################################################################

if __name__ == "__main__":
    # get participant list
    participant_dict = extract_participant_list(excel_file_path=PARTICIPANT_EXCEL_FILE, txt_file_path=PARTICIPANT_TXT_FILE, participant_dict=PARTICIPANTS)
    # validate participant list
    validate_participants(list(participant_dict.keys()))
    # assign gift matches
    secret_santa_draw_dict = assign_matches(list(participant_dict.keys()))
    # send emails
    email_secret_santa_draw(participant_dict, secret_santa_draw_dict, subject=EMAIL_SUBJECT, group_name=GROUP_NAME, instructions=INSTRUCTIONS, error_folder=ERROR_OUTPUT_FOLDER, sender_email="jesse.tonkin1999@gmail.com")
    
    
    
    
    
    
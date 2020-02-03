from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
from dotenv import load_dotenv


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

load_dotenv(dotenv_path='config.env')
# The ID of two spreadsheets.
mentor_sheet_id = os.getenv('MENTOR_SHEET_ID')
event_sheet_id = os.getenv('EVENT_SHEET_ID')

# Different ranges
event_range = os.getenv('EVENT_RANGE')
mentor_range = os.getenv('MENTOR_RANGE')

# cred
cred = os.getenv('GOOGLE_APPLICATION_CREDENTIALS')

# get google sheet service 
def get_service():
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = None
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('sheets', 'v4', credentials=creds)

# get google sheet
def get_sheet(service, sheet_id, sheet_range):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=sheet_range).execute()
    return result.get('values', [])

def populate_users_mentor(values, db):
    roles = db.collection(u'roles').where(u'value', u'==', u'Inductee').stream()
    role_id=None
    for doc in roles:
        role_id = doc.id
    for row in range(len(values)):
        # Look at the third, row[2], element of the row, which is the mentee's email 
        # and query for that user
        docs = db.collection(u'users').where(u'email', u'==', values[row][2]).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        for doc in docs:
            i += 1
        if i is 0:
            print(values[row][2], " is not in the database\n")
            # if the document doesn't exist, then populate it
            data = {
                u'email': values[row][2],
                u'mentorship': False,
                u'name': values[row][1],
                u'officer_signs': [],
                u'induction_points': 0,
                u'professional': False,
                u'role_id': role_id
            }
            db.collection(u'users').add(data)
            print(values[row][2], " is populated\n")
        elif i > 1:
            # This shouldn't happen, but just to check.
            print('More than one document for one email: ' + values[row][2] + '\n')
            return
        # else don't need to do anything



## ISSUE: Sheet column aignment had been updated. Change scripts to reflect the same.
## ISSUE: Find a way to verify emails in both columns before populating point.

def populate_users_event(values, db):
    roles = db.collection(u'roles').where(u'value', u'==', u'Inductee').stream()
    role_id=None
    for doc in roles:
        role_id = doc.id
    for row in range(len(values)):
        # Look at the third, row[2], element of the row, which is the mentee's email 
        # and query for that user
        docs = db.collection(u'users').where(u'email', u'==', values[row][1]).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        for doc in docs:
            i += 1
        if i is 0:
            print(values[row][1], " is not in the database\n")
            # if the document doesn't exist, then populate it
            data = {
                u'email': values[row][1],
                u'mentorship': False,
                u'name': values[row][2],
                u'officer_signs': [],
                u'induction_points': 0,
                u'professional': False,
                u'role_id': role_id
            }
            db.collection(u'users').add(data)
            print(values[row][1], " is populated\n")
        elif i > 1:
            # This shouldn't happen, but just to check.
            print('More than one document for one email: ' + values[row][1] + '\n')
        # else don't need to do anything

def update_event(values, db):
    point_reward_type = db.collection(u'pointRewardType').where(u'value', u'==', u'Induction Point').stream()
    reward_id=None
    for doc in point_reward_type:
        reward_id = doc.id
    user_dict = {}
    for row in range(len(values)):

        ## ISSUE: Queries for user doc multiple times.

        docs = db.collection(u'users').where(u'email', u'==', values[row][1]).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        doc_id = None
        for doc in docs:
            i += 1
            doc_id = doc.id

        if i is not 1:
            # This shouldn't happen but just to check
            print(values[row][1] + "More or less than one doc is returned\n")
            return

            ## ISSUE: Breaks on error AFTER populating some values.
            ## ISSUE: Return row number so and exit so that next time the script runs from this row onwards

        else:
            # Now we can udate the events
            date_time = datetime.datetime.strptime(values[row][0], "%m/%d/%Y %H:%M:%S")
           
            data = {
                u'created': date_time,
                u'event_name': values[row][7],
                u'officer_name': values[row][9],
                u'pointrewardtype_id': reward_id,
                u'user_id': doc_id,
                u'value': float(values[row][8])
            }

            db.collection(u'pointReward').add(data)

            print("Finished updating events\n")

            if doc_id in user_dict:
                user_dict[doc_id]['induction_points'] += float(values[row][8])
                user_dict[doc_id]['officer_signs'].add(values[row][9])
            else:
                user_dict[doc_id] = {'induction_points': float(values[row][8]), 'officer_signs': {values[row][9]}}
            
            if 'interview' in values[row][7].lower() or 'resume' in values[row][7].lower():
                user_dict[doc_id]['professional'] = True
        
    for key, value in user_dict.items():
        value['induction_points'] = firestore.Increment(value['induction_points'])
        value['officer_signs'] = firestore.ArrayUnion(list(value['officer_signs']))
        db.collection(u'users').document(key).update(value)
       

def update_mentor_event(values, db):
    point_reward_type = db.collection(u'pointRewardType').where(u'value', u'==', u'Induction Point').stream()
    reward_id = None
    for doc in point_reward_type:
        reward_id = doc.id
    user_dict = {}
    for row in range(len(values)):

        ## ISSUE: Queries for user doc multiple times.

        docs = db.collection(u'users').where(u'email', u'==', values[row][2]).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        doc_id = None
        for doc in docs:
            i += 1
            doc_id = doc.id

        if i is not 1:
            # This shouldn't happen but just to check
            print(values[row][2] + "More or less than one doc is returned\n")
            return

            ## ISSUE: Breaks on error AFTER populating some values.
            ## ISSUE: Return row number so and exit so that next time the script runs from this row onwards

        else:
            # Now we can udate the events
            date_time = datetime.datetime.strptime(values[row][0], "%m/%d/%Y %H:%M:%S")
            data = {
                u'created': date_time,
                u'event_name': u'Mentor 1:1',
                u'officer_name': values[row][3],
                u'pointrewardtype_id': reward_id,
                u'user_id': doc_id,
                u'value': 1
            }

            db.collection(u'pointReward').add(data)
            
            print("Finished updating mentor events\n")
            
            if doc_id in user_dict:
                user_dict[doc_id]['induction_points'] += 1
                user_dict[doc_id]['officer_signs'].add(values[row][3])
            else:
                user_dict[doc_id] = {'induction_points': 1, 'officer_signs': {values[row][3]}}
            
            user_dict[doc_id]['mentorship'] = True
    
    for key, value in user_dict.items():
        value['induction_points'] = firestore.Increment(value['induction_points'])
        value['officer_signs'] = firestore.ArrayUnion(list(value['officer_signs']))
        db.collection(u'users').document(key).update(value)
            
        
def main():
    # get the google sheet service
    service = get_service()
    # get all the values into a 2d array 
    values_mentor = get_sheet(service, mentor_sheet_id, mentor_range)
    values_event = get_sheet(service, event_sheet_id, event_range)

    # start firebase and get access to the databse
    firebase_admin.initialize_app()
    db = firestore.client()  
    

    ## ISSUE: Change update_** method calls to store retunred row value and write that down in the config file.
    if(len(values_mentor) > 0):
        populate_users_mentor(values_mentor, db)
        update_mentor_event(values_mentor, db)
    
    
    if(len(values_event) > 0):
        populate_users_event(values_event, db)
        update_event(values_event, db)
    

    # This is the only crazy and stupid way I can think of to rewrite the env file
    # If you found any other way please use them.
    # This way works if you have exactly these 5 fields.
    starting_index = event_range.replace('Sheet1!A','')
    starting_index = starting_index.replace(':K', '')
    starting_index_mentor = mentor_range.replace('Sheet1!A','')
    starting_index_mentor = starting_index_mentor.replace(':G', '')

    ## ISSUE: Update range value that gets written down to config file.
    lines = ['EVENT_SHEET_ID="'+event_sheet_id+'"\n', 'MENTOR_SHEET_ID="'+mentor_sheet_id+'"\n',
            'GOOGLE_APPLICATION_CREDENTIALS="'+cred+'"\n', 
            'EVENT_RANGE="Sheet1!A'+str(len(values_event)+int(starting_index))+':K"\n', 
            'MENTOR_RANGE="Sheet1!A'+str(len(values_mentor)+int(starting_index_mentor))+':G"\n']
    with open('config.env', 'w') as f:
        f.writelines(lines)


if __name__ == '__main__':
    main()
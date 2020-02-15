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
import sys


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

load_dotenv(dotenv_path='config.env')
# The ID of two spreadsheets.
mentor_sheet_id = os.getenv('MENTOR_SHEET_ID')
event_sheet_id = os.getenv('EVENT_SHEET_ID')

# Different ranges
event_range = os.getenv('EVENT_RANGE')
mentor_range = os.getenv('MENTOR_RANGE')

# Starting row number
event_start = int(event_range.replace('Sheet1!A','').replace(':L', ''))
mentor_start = int(mentor_range.replace('Sheet1!A', '').replace(':I', ''))

DRY_RUN = 0
MODIFY = 1
mode = DRY_RUN

users_docID = {}
users_data = {}
roles = {}

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

def populate_users(sheetName, values, startRow, nameCol, emailCol, confirmEmailCol, db):
    error = False
    roles = db.collection(u'roles').where(u'value', u'==', u'Inductee').stream()
    role_id=None

    num_docs = 0

    for doc in roles:
        role_id = doc.id
        num_docs += 1
    
    if(num_docs != 1):
        print("Multiple documents for enum Inductee")
        return True

    for row in range(len(values)):
        # Look at the third, row[2], element of the row, which is the mentee's email 
        # and query for that user

        if values[row][emailCol].lower() != values[row][confirmEmailCol].lower() and (len(values[row][confirmEmailCol].lower()) != 0):
            print("Mismatched emails in " + sheetName + " form at row: " + str(startRow + row))
            error = True
            continue

        user_email = values[row][emailCol].lower().strip()

        if user_email in users_docID:
            print("User known to exist")
            continue

        userDocID = None

        docs = db.collection(u'users').where(u'email', u'==', user_email).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        for doc in docs:
            i += 1
            userDocID = doc.id
        if i is 0:
            print(user_email, " is not in the database\n")
            # if the document doesn't exist, then populate it
            data = {
                u'email': user_email,
                u'mentorship': False,
                u'name': values[row][nameCol],
                u'officer_signs': [],
                u'induction_points': 0,
                u'professional': False,
                u'role_id': role_id
            }

            if mode == MODIFY:
                userDocID = db.collection(u'users').add(data)[1].id
            elif mode == DRY_RUN:
                print("Creating user document in the database for email " + user_email)
            
            print(user_email, " is populated\n")
        elif i > 1:
            # This shouldn't happen, but just to check.
            userDocID = None
            print('More than one document for one email: ' + values[row][emailCol] + '\n')
            error = True
        # else don't need to do anything
        
        if userDocID != None:
            print("Added email to global list of emails")
            users_docID[user_email] = userDocID

    return error

# Method replaced by populate_users
def populate_users_mentor(values, db):
    error = False
    roles = db.collection(u'roles').where(u'value', u'==', u'Inductee').stream()
    role_id=None

    num_docs = 0

    for doc in roles:
        role_id = doc.id
        num_docs += 1
    
    if(num_docs != 1):
        print("Multiple documents for enum Inductee")
        return True

    for row in range(len(values)):
        # Look at the third, row[2], element of the row, which is the mentee's email 
        # and query for that user

        if values[row][2].lower() != values[row][3].lower() and (len(values[row][3].lower()) != 0):
            print("Mismatched emails in mentor form at row: " + str(mentor_start + row))
            # print("Mismatched emails in mentor form at row: {}".format(mentor_start + row))
            error = True
            continue

        user_email = values[row][2].lower().strip()

        if user_email in users_docID:
            print("User known to exist")
            continue

        userDocID = None

        docs = db.collection(u'users').where(u'email', u'==', user_email).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        for doc in docs:
            i += 1
            userDocID = doc.id
        if i is 0:
            print(user_email, " is not in the database\n")
            # if the document doesn't exist, then populate it
            data = {
                u'email': user_email,
                u'mentorship': False,
                u'name': values[row][1],
                u'officer_signs': [],
                u'induction_points': 0,
                u'professional': False,
                u'role_id': role_id
            }

            if mode == MODIFY:
                userDocID = db.collection(u'users').add(data)[1].id
            elif mode == DRY_RUN:
                print("Creating user document in the database for email " + user_email)
            
            print(user_email, " is populated\n")
        elif i > 1:
            # This shouldn't happen, but just to check.
            userDocID = None
            print('More than one document for one email: ' + values[row][2] + '\n')
            error = True
        # else don't need to do anything
        
        if userDocID != None:
            print("Added email to global list of emails")
            users_docID[user_email] = userDocID

    return error

## ISSUE: Sheet column aignment had been updated. Change scripts to reflect the same.
## ISSUE: Find a way to verify emails in both columns before populating point.

# Method replaced by populate_users
def populate_users_event(values, db):
    error = False
    roles = db.collection(u'roles').where(u'value', u'==', u'Inductee').stream()
    role_id=None

    num_docs = 0

    for doc in roles:
        role_id = doc.id
        num_docs += 1
    
    if(num_docs != 1):
        print("Multiple documents for enum Inductee")
        return True

    for row in range(len(values)):

        if (values[row][1].lower() != values[row][2].lower()) and (len(values[row][2].lower()) != 0):
            print("Mismatched emails in event form at row: {}".format(event_start + row))
            error = True
            continue

        user_email = values[row][1].lower().strip()

        if user_email in users_docID:
            print("User known to exist")
            continue

        userDocID = None

        # Look at the third, row[2], element of the row, which is the mentee's email 
        # and query for that user
        docs = db.collection(u'users').where(u'email', u'==', user_email).stream()
        # Use i to count how many documents are returned from the query
        i = 0
        for doc in docs:
            i += 1
            userDocID = doc.id
        if i is 0:
            print(user_email, " is not in the database\n")
            # if the document doesn't exist, then populate it
            data = {
                u'email': user_email,
                u'mentorship': False,
                u'name': values[row][3],
                u'officer_signs': [],
                u'induction_points': 0,
                u'professional': False,
                u'role_id': role_id
            }

            if mode == MODIFY:
                userDocID = db.collection(u'users').add(data)[1].id
            elif mode == DRY_RUN:
                print("Creating user document in the database for email " + user_email)

            print(user_email, " is populated\n")
        elif i > 1:
            # This shouldn't happen, but just to check.
            userDocID = None
            print('More than one document for one email: ' + values[row][1] + '\n')
            error = True
        # else don't need to do anything

        if userDocID != None:
            print("Added email to global list of emails")
            users_docID[user_email] = userDocID
    
    return error

def update_event(values, db):
    point_reward_type = db.collection(u'pointRewardType').where(u'value', u'==', u'Induction Point').stream()
    
    reward_id=None
    
    for doc in point_reward_type:
        reward_id = doc.id
    
    for row in range(len(values)):

        if (values[row][1].lower() != values[row][2].lower()) and (len(values[row][2].lower()) != 0):
            print("Mismatched emails in event form at row: {}".format(event_start + row))
            return row

        ## ISSUE: Queries for user doc multiple times.

        user_email = values[row][1].lower().strip()

        userDocID = users_docID.get(user_email, None)

        if userDocID == None:
            print("User document for user with email " + user_email + " does not exist.")
            continue
        
        # Now we can udate the events
        date_time = datetime.datetime.strptime(values[row][0], "%m/%d/%Y %H:%M:%S")
        
        data = {
            u'created': date_time,
            u'event_name': values[row][8],
            u'officer_name': values[row][10],
            u'pointrewardtype_id': reward_id,
            u'user_id': userDocID,
            u'value': float(values[row][9])
        }

        if mode == MODIFY:
            db.collection(u'pointReward').add(data)
        elif mode == DRY_RUN:
            print("\nAdding the following point to the database:")
            print(data)

        print("Finished updating events\n")

        if userDocID in users_data:
            users_data[userDocID]['induction_points'] += float(values[row][9])
            users_data[userDocID]['officer_signs'].add(values[row][10])
        else:
            users_data[userDocID] = {'induction_points': float(values[row][9]), 'officer_signs': {values[row][10]}}
        
        if 'interview' in values[row][8].lower() or 'resume' in values[row][8].lower():
            users_data[userDocID]['professional'] = True
    
    return len(values)

def update_mentor_event(values, db):
    point_reward_type = db.collection(u'pointRewardType').where(u'value', u'==', u'Induction Point').stream()
    
    reward_id = None
    
    for doc in point_reward_type:
        reward_id = doc.id
    
    for row in range(len(values)):

        if values[row][2].lower() != values[row][3].lower() and (len(values[row][3].lower()) != 0):
            print("Mismatched emails in mentor form at row: {}".format(mentor_start + row))
            return row
        
        user_email = values[row][2].lower().strip()

        ## ISSUE: Queries for user doc multiple times

        userDocID = users_docID.get(user_email, None)

        if userDocID == None:
            print("User document for user with email " + user_email + " does not exist.")
            continue
        
        date_time = datetime.datetime.strptime(values[row][0], "%m/%d/%Y %H:%M:%S")
        data = {
            u'created': date_time,
            u'event_name': u'Mentor 1:1',
            u'officer_name': values[row][4],
            u'pointrewardtype_id': reward_id,
            u'user_id': userDocID,
            u'value': 1
        }

        if mode == MODIFY:
            db.collection(u'pointReward').add(data)
        elif mode == DRY_RUN:
            print("\nAdding the following point to the database:")
            print(data)
        
        print("Finished updating mentor events\n")
        
        if userDocID in users_data:
            users_data[userDocID]['induction_points'] += 1
            users_data[userDocID]['officer_signs'].add(values[row][4])
        else:
            users_data[userDocID] = {'induction_points': 1, 'officer_signs': {values[row][4]}}
        
        users_data[userDocID]['mentorship'] = True

    return len(values)

def getEnumMap(collectionName, db):
    roles = {}
    rolesDocs = db.collection(collectionName).stream()

    for doc in rolesDocs:
        roles[doc.get('value')] = doc.id
    
    return roles

def main():

    global mode
    global users_docID
    global users_data
    global roles

    print("Mode before processing input")
    print(mode)

    if(len(sys.argv) != 2):
        print("Please use -D to indicate a dry-run of the program and -M for the actual run which will modify the database.")
        return
    elif (sys.argv[1] == "-D"):
        print("Executing program in dry-run mode. Database will not be modified.")
        mode = DRY_RUN
    elif (sys.argv[1] == "-M"):
        print("Executing program in modify mode. Database will be modified.")
        mode = MODIFY
    else:
        print("Please use -D to indicate a dry-run of the program and -M for the actual run which will modify the database.")
        return

    print("Mode after processing input")
    print(mode)
    # return
    # get the google sheet service
    service = get_service()
    # get all the values into a 2d array 
    values_mentor = get_sheet(service, mentor_sheet_id, mentor_range)
    values_event = get_sheet(service, event_sheet_id, event_range)

    # start firebase and get access to the databse
    firebase_admin.initialize_app()
    db = firestore.client()  
    
    roles = getEnumMap("roles", db)

    print(roles)
    return
    print(event_start)
    print(mentor_start)

    processed_mentor_pts = 0
    processed_events_pts = 0

    ## ISSUE: Change update_** method calls to store retunred row value and write that down in the config file.

    print("\n\nWORKING WITH MENTOR POINTS SHEET\n\n")
    # if(len(values_mentor) > 0 and not populate_users_mentor(values_mentor, db)):
    if(len(values_mentor) > 0 and not populate_users("mentor", values_mentor, mentor_start, 1, 2, 3, db) and not populate_users("mentor", values_mentor, mentor_start, 4, 5, 6, db)):
        print("\n\nPROCESSING MENTOR POINTS SHEET\n\n")
        processed_mentor_pts = update_mentor_event(values_mentor, db)
    
    
    print("\n\nWORKING WITH EVENT POINTS SHEET\n\n")
    # if(len(values_event) > 0 and not populate_users_event(values_event, db)):
    if(len(values_event) > 0 and not populate_users("event", values_event, event_start, 3, 1, 2, db)):
        print("\n\nPROCESSING EVENT POINTS SHEET\n\n")
        processed_events_pts = update_event(values_event, db)
    

    for key, value in users_data.items():
        value['induction_points'] = firestore.Increment(value['induction_points'])
        value['officer_signs'] = firestore.ArrayUnion(list(value['officer_signs']))
        if mode == MODIFY:
            db.collection(u'users').document(key).update(value)
        elif mode == DRY_RUN:
            print("\nUpdating document with id " + key + " with following data:")
            print(value)

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
            'EVENT_RANGE="Sheet1!A'+str(processed_events_pts+int(event_start))+':L"\n', 
            'MENTOR_RANGE="Sheet1!A'+str(processed_mentor_pts+int(mentor_start))+':I"\n']
    with open('config.env', 'w') as f:
        f.writelines(lines)


if __name__ == '__main__':
    main()
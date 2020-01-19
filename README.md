# Points_Transfer
This project is a script that transfers all the data from two google sheets, one sheet recording all the regular events 
and the other sheet recording all the Mentor 1:1 events, to cloud firestore.
# Setup google sheet API
https://developers.google.com/sheets/api/quickstart/python
There is detailed instructions on how to setup google sheet api on the website.
Here are the basic steps:
1. Click "Enable the Google Sheet API" on the website and download the credentials.json file to your working directory.
2. run 
```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dotenv
```
# Setup cloud firestore
https://firebase.google.com/docs/admin/setup/
There is detailed instructions on how to setup firebase-admin on this website.
Here are the basic steps:
1. run
```
pip install --user firebase-admin
```
2. generate a private key file for your account and set the GOOGLE_APPLICATION_CREDENTIALS environment variable in your config.env file.
# Setup the config.env file
There should be five fields in the file

|                 Key                  |                Description                                   |
|--------------------------------------|--------------------------------------------------------------|
|        EVENT_SHEET_ID                | Google Sheet ID of Event Point Log                           |
|        EVENT_RANGE                   | Google Sheet Range of Event Point Log                        |
|        MENTOR_SHEET_ID               | Google Sheet ID of Mentor Point Log                          |
|        MENTOR_RANGE                  | Google Sheet Range of Mentor Point Log                       |
| GOOGLE_APPLICATION_CREDENTIALS       | private key generated from firebase console                  |

Note: the range of the two google sheets defines where the program should read.\
For event range, it should look something like, Sheet1!A\<row to start reading\>:K.\
For mentor range, it should look something like, Sheet1!A\<row to start reading\>:G.

# Clone the repo to your local working directory.


from __future__ import print_function
import webbrowser
import pandas as pd
import time
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

# import datetime
from datetime import datetime
from datetime import timedelta
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


def main():

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'creds.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    meetings = pd.DataFrame(columns=['Name', 'Link', 'Time'])
    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        week_later = (datetime.now() + timedelta(weeks=1)).isoformat() + 'Z'
        print('Scheduling automatic zoom meeting joins for the next week.')
        events_result = service.events().list(calendarId='primary', timeMin=now,
                                              timeMax=week_later, singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])
        for event in events:
            if ('zoom' in event["location"].lower() or 'teams' in event["location"].lower()):
                # event_series = pd.Series(event, index=)
                event_to_schedule = pd.Series([event["summary"], event["location"], event["start"]])
                meetings = pd.concat([meetings, event_to_schedule.to_frame().T])
            
        if meetings.empty is True:
            print("No remote meetings found for the next week.")
        
        if not events_result:
            print("No upcoming events found.")
            return

        # Prints the start and name of the next 10 events
        # for event in events:
        #     start = event['start'].get('dateTime', event['start'].get('date'))
        #     # print(start, event['summary'])

    except HttpError as error:
        print('An error occurred: %s' % error)
    
    for event in events:
            start = event['start'].get('dateTime', 'date')
            #print(start, event['summary'])

    

    while True:
        start = datetime.now()
        end = start + timedelta(minutes=5)
        next_event = events[0]
        next_event_start_time = next_event['start']['dateTime']
        next_event_start_time = next_event_start_time[:-3] + '' + next_event_start_time[-2:] # remove colon from timezone
        next_event_start_time = next_event_start_time.replace("T", '/')
        print("next event start time: ", next_event_start_time)

        if start <= datetime.strptime(next_event_start_time, '%Y-%m-%d/%H:%M:%S.%f-%z') <= end:
            join_live_meeting(next_event.get('location'))
            # event_to_schedule.drop(index=event_to_schedule.index[0], axis=0, inplace=True)
            events.pop(0)
        else:
            print("attempting to join upcoming meeting, but it will not start soon.")
        time.sleep(60)

    
    
def join_live_meeting(url):
    driver = webdriver.Firefox()
    driver.implicitly_wait(5) # seconds
    driver.get(url)

def view_recording(url, passcode): 
    driver = webdriver.Firefox()
    driver.implicitly_wait(5) # seconds
    driver.get(url)
    element = driver.find_element(By.ID, 'passcode').send_keys(passcode)

    loc = None
    while loc == None:
        loc = pyautogui.locateCenterOnScreen('watch_recording_button.png', confidence=0.9)
    print("center:", loc)
    pyautogui.click(loc)

if __name__ == '__main__':
    main()
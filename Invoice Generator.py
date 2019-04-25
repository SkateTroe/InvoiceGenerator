from __future__ import print_function
import datetime
from datetime import date, timezone
from datetime import timedelta
from dateutil.parser import parse
import time
import pickle
import os.path
import xlwt
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

class clEvent:
	def __init__(self, date, duration, description, price):
		self.date = date
		self.duration = duration.total_seconds() / 3600
		self.description = description
		self.price = int(price.strip('$'))
		self.total = self.duration * self.price

def main():
    """Accesses Google Calendar API to pull information from all Tutoring appointments
    	for the current week beginning on Friday.
    """

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API--this is the ID string for my Tutoring Appointments calendar.
    strCalID = "mst.edu_86nbm4nc7464vvt6qhpoffmtl8@group.calendar.google.com"
    #individual functions to calculate the date of the previous Friday and the upcoming Thursday
    dtFriday = fnFriday()
    dtThursday = fnThursday()

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time

    #Gathers all events between dtFriday (last Friday) and dtThursday (this Thursday) in events_result
    events_result = service.events().list(calendarId=strCalID, timeMin=dtFriday,
                                        timeMax=dtThursday, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    lst = []

    if not events:
        print('No appointments for this week.')
    for event in events:

    	#parse() turns string into dateTime, allowing date math later on.
    	#parse().strftime shifts from string to datetime back to string, but with proper formatting
        dtstart=parse(event['start'].get('dateTime'))
        dtend= parse(event['end'].get('dateTime'))
        #using parse() and strftime() to render the date in an easy-to-read MM/DD/YY format
        dtDate=parse(event['start'].get('dateTime')).strftime('%m/%d/%y')

        #Add each event as a clEvent object to be output to Excel
        lst.append(clEvent(dtDate,dtend - dtstart,event['summary'], event['description']))
    #Don't know if this should or shouldn't be a separate function, but I like it.
    fnOutput(lst, dtThursday, dtFriday)

def fnThursday():
	today = datetime.datetime.now(timezone.utc).astimezone()
	#Thursday is ID'd as '3'
	offset = 3 - today.weekday()
	#Mathematical Magic.
	dtThursday = today + timedelta(days = offset) if offset >= 0 else today - timedelta(days = offset + 7)
	#ensuring we include all of Thursday.
	dtThursday = dtThursday.replace(hour=23, minute=59, second=59)
	print('Ends: ' + dtThursday.isoformat('T'))
	return(dtThursday.isoformat('T'))

def fnFriday():
	today = datetime.datetime.now(timezone.utc).astimezone()
	offset = today.weekday() - 4
	dtFriday = today - timedelta(days = offset) if offset >= 0 else today - timedelta(days = offset + 7)
	#ensuring we include none of Friday--that's for next week.
	dtFriday = dtFriday.replace(hour=00, minute=00, second=00)
	print('Starts: ' + dtFriday.isoformat('T'))
	return(dtFriday.isoformat('T'))

def fnOutput(lst, dtThursday, dtFriday):
	#using xlwt, we create the workbook and add a sheet.
	book = xlwt.Workbook()
	sheet1 = book.add_sheet("shData")
	#The invoice date is the Friday after the pay period. So we take the day after dtThursday.
	dtInvoiceDate = parse(dtThursday) + timedelta(days = 1)

	#Setting a couple of useful dates for Word to pull from.
	sheet1.row(1).write(5, parse(dtFriday).strftime('%m/%d/%y'))
	sheet1.row(1).write(6, parse(dtThursday).strftime('%m/%d/%y'))
	sheet1.row(1).write(7, dtInvoiceDate.strftime('%B %d, %Y'))
	#Column Headings
	sheet1.row(0).write(0, "StartDate")
	sheet1.row(0).write(1, "Duration")
	sheet1.row(0).write(2, "Description")
	sheet1.row(0).write(3, "HourlyPrice")
	sheet1.row(0).write(4, "TotalPrice")

	#For each clEvent in lst, fill out a row in Excel
	for i in range (0, len(lst)):
		row = sheet1.row(i+1)
		row.write(0, lst[i].date)
		row.write(1, lst[i].duration)
		row.write(2, lst[i].description)
		row.write(3, str(lst[i].price))
		row.write(4, str(lst[i].total))


	#WWJD?
	book.save("dbCalendarData.xls")

if __name__ == '__main__':
    main()
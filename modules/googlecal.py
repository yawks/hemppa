from __future__ import print_function

import os
import os.path
import pickle
from datetime import datetime, timedelta
import re
import pytz

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from modules.common.module import BotModule
from modules.common.exceptions import UploadFailed
#
# Google calendar notifications
#
# Note: Provide a token.pickle file for the service.
# It's created on first run (run from console!) and
# can be copied to another computer.

VIDEOCALL_DESC_REGEXP = [
    r"href=\"(https:..primetime.bluejeans.com.a2m.live-event.([^\/\"])*\")",
    r"(https://(\w+\.)?zoom.\w{,2}(/j)?/\d+\??[^\"\n\s]*)",
    r"(https://meet.google.com/[^\n]*)",
    r"(https://meet.lync.com/[^\n]*)",
    r"(https://meet.office.com/[^\n]*)",
    r"(https://meet.microsoft.com/[^\n]*)",
    r"(https://teams.microsoft.com/[^>\n]*)",
]


class MatrixModule(BotModule):
    def __init__(self, name):
        super().__init__(name)
        self.credentials_file = "credentials.json"
        self.SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
        self.bot = None
        self.service = None
        self.service_name = 'googlecal'
        self.calendar_rooms = dict()  # Contains room_id -> [calid, calid] ..
        self.enabled = True
        self.poll_interval_min = 1
        self.owner_only = True
        self.send_all = True

    def matrix_start(self, bot):
        super().matrix_start(bot)
        self.bot = bot
        creds = None

        if not os.path.exists(self.credentials_file) or os.path.getsize(self.credentials_file) == 0:
            return  # No-op if not set up

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
                self.logger.info('Loaded existing pickle file')
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            self.logger.warn('No credentials or credentials not valid!')
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, self.SCOPES)
                # urn:ietf:wg:oauth:2.0:oob
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
                self.logger.info('Pickle saved')

        self.service = build('calendar', 'v3', credentials=creds)

        try:
            calendar_list = self.service.calendarList().list().execute()['items']
            self.logger.info(f'Google calendar set up successfully with access to {len(calendar_list)} calendars:\n')
            for calendar in calendar_list:
                self.logger.info(f"{calendar['summary']} - + {calendar['id']}")
        except Exception:
            self.logger.error('Getting calendar list failed!')

    async def matrix_message(self, bot, room, event):
        if not self.service:
            await bot.send_text(room, 'Google calendar not set up for this bot.')
            return
        args = event.body.split()
        events = []
        calendars = self.calendar_rooms.get(room.room_id) or []

        if len(args) == 2:
            if args[1] == 'today':
                for calid in calendars:
                    self.logger.info(f'Listing events in cal {calid}')
                    events = events + self.list_today(calid)
            elif args[1] == 'list':
                await bot.send_text(room, 'Calendars in this room: ' + str(self.calendar_rooms.get(room.room_id)))
                return
            elif args[1] == 'listavailable':
                calendar_list = self.service.calendarList().list().execute()['items']
                calendars = 'Available calendars: \n'
                for calendar in calendar_list:
                    if calendar['summary'] not in self.calendar_rooms.get(room.room_id, []):
                        calendars += f" - {calendar['summary']}\n"
                await bot.send_text(room, calendars)
                return

        elif len(args) == 3:
            if args[1] == 'add':
                bot.must_be_admin(room, event)

                calid = args[2]
                self.logger.info(f'Adding calendar {calid} to room id {room.room_id}')

                if self.calendar_rooms.get(room.room_id):
                    if calid not in self.calendar_rooms[room.room_id]:
                        self.calendar_rooms[room.room_id].append(calid)
                    else:
                        await bot.send_text(room, 'This google calendar already added in this room!')
                        return
                else:
                    self.calendar_rooms[room.room_id] = [calid]

                self.logger.info(f'Calendars now for this room {self.calendar_rooms.get(room.room_id)}')

                bot.save_settings()

                await bot.send_text(room, 'Added new google calendar to this room')
                return

            if args[1] == 'del':
                bot.must_be_admin(room, event)

                calid = args[2]
                self.logger.info(f'Removing calendar {calid} from room id {room.room_id}')

                if self.calendar_rooms.get(room.room_id):
                    self.calendar_rooms[room.room_id].remove(calid)

                self.logger.info(f'Calendars now for this room {self.calendar_rooms.get(room.room_id)}')

                bot.save_settings()

                await bot.send_text(room, 'Removed google calendar from this room')
                return

        else:
            for calid in calendars:
                self.logger.info(f'Listing events in cal {calid}')
                events = events + self.list_upcoming(calid)

        if len(events) > 0:
            self.logger.info(f'Found {len(events)} events')
            await self.send_events(bot, events, room)
        else:
            self.logger.info(f'No events found')
            await bot.send_text(room, 'No events found, try again later :)')

    async def send_events(self, bot, events, room):
        previous_day = None
        events_of_same_day = []
        for event in events:
            start_date = self.parse_date(event['start'].get('dateTime', event['start'].get('date')))
            current_day = datetime.strftime(start_date, '%a %d %b')
            if previous_day is None or current_day == previous_day:
                events_of_same_day.append(event)
            else:
                await self.send_html_same_day_events(bot, room, events_of_same_day, previous_day)
                events_of_same_day = []

            previous_day = current_day

        await self.send_html_same_day_events(bot, room, events_of_same_day, previous_day)

    async def send_html_same_day_events(self, bot, room, events_of_same_day, current_day):
        html = f"<hr/><h2>📅 {current_day}</h2>\n"
        text = f" 📅 {current_day}\n"
        for event in events_of_same_day:
            start_hour, end_hour = self.get_event_hours(event)

            img_videocall = await get_videocall_logo_from_summary(bot, event)
            html += f'<strong>{start_hour}{end_hour}</strong> <a href="{event["htmlLink"]}">{event["summary"]}</a> {img_videocall}<br/>\n'
            text += f' - {start_hour}{end_hour} \n {event["summary"]}\n\n'
        await bot.send_html(room, html, text)

    def get_event_hours(self, event):
        start_hour = "All day"
        end_hour = ""
        if event['start'].get('dateTime', event['start'].get('date')) != event['end'].get('dateTime', event['end'].get('date')):
            start_hour = self.reformat_strdate(event['start'].get('dateTime', event['start'].get('date')))
            if start_hour != "All day":
                end_hour = " - " + self.reformat_strdate(event['end'].get('dateTime', event['end'].get('date')))
        return start_hour, end_hour

    def list_upcoming(self, calid):
        start_time = datetime.utcnow()
        now = start_time.isoformat() + 'Z'
        events_result = self.service.events().list(calendarId=calid, timeMin=now,
                                                   maxResults=10, singleEvents=True,
                                                   orderBy='startTime').execute()
        events = events_result.get('items', [])
        return events

    def list_today(self, calid):
        start_time = datetime.utcnow()
        end_time = start_time + timedelta(hours=48)
        now = start_time.isoformat() + 'Z'
        end = end_time.isoformat() + 'Z'
        self.logger.info(f'Looking for events between {now} {end}')
        events_result = self.service.events().list(calendarId=calid, timeMin=now,
                                                   timeMax=end, maxResults=10, singleEvents=True,
                                                   orderBy='startTime').execute()
        return events_result.get('items', [])

    def help(self):
        return 'Google calendar. Lists 10 next events by default. today = list today\'s events.'

    def get_settings(self):
        data = super().get_settings()
        data['calendar_rooms'] = self.calendar_rooms
        return data

    def set_settings(self, data):
        super().set_settings(data)
        if data.get('calendar_rooms'):
            self.calendar_rooms = data['calendar_rooms']

    def parse_date(self, start: str) -> datetime:
        try:
            dt = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S%z')
        except ValueError:
            dt = datetime.strptime(start, '%Y-%m-%d')

        return dt

    def reformat_strdate(self, start):
        try:
            dt = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S%z')
            return dt.strftime("%H:%M")
        except ValueError:
            return "All day"

    async def matrix_poll(self, bot, pollcount):
        if pollcount % 6 == 0:  # every minute
            for room_id in self.calendar_rooms:
                calendars = self.calendar_rooms.get(room_id) or []
                for calid in calendars:
                    start_time = datetime.utcnow()
                    start_time = start_time + timedelta(minutes=2)
                    events_result = self.service.events().list(calendarId=calid, timeMin=start_time.isoformat() + 'Z',
                                                               maxResults=10, singleEvents=True,
                                                               orderBy='startTime').execute()
                    for event in events_result.get('items', []):
                        event_start_time = self.parse_date(event['start'].get('dateTime', event['start'].get('date')))
                        if (event_start_time.tzinfo is not None and event_start_time <= start_time.replace(tzinfo=pytz.utc)) or \
                                (event_start_time.tzinfo is None and event_start_time <= start_time):
                            start_hour, end_hour = self.get_event_hours(event)
                            html, text = self.get_html_and_text_messages(event, start_hour, end_hour)
                            await bot.send_html_with_room_id(room_id, html, text)

    def get_html_and_text_messages(self, event, start_hour, end_hour):
        html = f'<hr/>📣 <i>1 minute until this event:</i><br/><strong>{start_hour}{end_hour} <a href="{event["htmlLink"]}">{event["summary"]}</a></strong><br/>\n'
        if 'location' in event:
            html += '<br/>🌍 ' + event['location']
        if 'attendees' in event:
            html += '<br/>🙋 '
            cpt = 0
            for attendee in event.get('attendees', []):
                html += f'{attendee.get("displayName", attendee["email"])}' + (" <i>(Organizer)</i>" if event.get('organizer', {}).get('email', '') == attendee['email'] else '')
                if cpt < len(event['attendees']) - 1:
                    html += ', '
                cpt += 1
        if event.get('description', '').strip() != "":
            html += '<br/>----------------<br/>' + event['description']

        text = f' - ** {start_hour}{end_hour}** \n {event["summary"]}\n\n'
        return html, text


async def get_videocall_logo_from_summary(bot, event) -> str:
    img_html = ''
    url = event.get('conferenceData', {}).get(
        'entryPoints', [{}])[0].get('uri', '')
    if url == '':
        for reg in VIDEOCALL_DESC_REGEXP:
            match = re.search(reg, event.get('description', ''))
            if match:
                url = match.groups()[0]
                break

    matrix_uri = ''
    try:
        if url.find('meet.google.com') > -1:
            matrix_uri, _, _, _, _ = await bot.upload_image('https://upload.wikimedia.org/wikipedia/commons/thumb/9/9b/Google_Meet_icon_%282020%29.svg/12px-Google_Meet_icon_%282020%29.svg.png?20221213135236', blob_content_type="image/png")
        elif url.find('teams.microsoft.com') > -1:
            matrix_uri, _, _, _, _ = await bot.upload_image('https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg/12px-Microsoft_Office_Teams_%282018%E2%80%93present%29.svg.png?20210603103011', blob_content_type="image/png")
    except (UploadFailed, TypeError, ValueError):
        print(f"Something went wrong uploading meet logo.")

    if matrix_uri != '':
        img_html = f'<img src="{matrix_uri}"/>'

    return img_html

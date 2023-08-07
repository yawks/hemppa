from __future__ import print_function
from typing import Dict, List, Optional, Tuple
import os.path
import timeago
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from modules.common.module import BotModule
from modules.common.exceptions import UploadFailed
#
# Google task notifications
#


SCOPES = ['https://www.googleapis.com/auth/tasks']


class MatrixModule(BotModule):

    def __init__(self, name):
        super().__init__(name)
        self.service_name = 'googletasks'
        self.tasklists_rooms = {}
        self.enabled = True
        #self.service = None
        self.poll_interval_min = 1

    def matrix_start(self, bot):
        super().matrix_start(bot)
        self.add_module_aliases(bot, ['newname', 'anothername'])

    def get_settings(self):
        data = super().get_settings()
        data['tasklists_rooms'] = self.tasklists_rooms
        return data

    def set_settings(self, data):
        super().set_settings(data)
        if data.get('tasklists_rooms'):
            self.tasklists_rooms = data['tasklists_rooms']

    async def matrix_message(self, bot, room, event):
        if not self.service:
            await bot.send_text(room, 'Google calendar not set up for this bot.')
            return
        args = event.body.split()
        events = []
        tasklists = self.tasklists_rooms.get(room.room_id) or []
        
        if len(args) == 2:
            if args[1] == 'today':
                for tasklistid in tasklists:
                    self.logger.info(f'Listing events in tasklist {tasklistid}')
                    events = events + self.list_today(tasklistid)
                    group_in_date = datetime.now().strftime("%a %d %b")  # force to group all events in the same day (case of events on multiple days)
            elif args[1] == 'list':
                await bot.send_text(room, 'Calendars in this room: ' + str(self.calendar_rooms.get(room.room_id)))
                return
            elif args[1] == 'listavailable':
                calendar_list = self.service.calendarList().list().execute()['items']
                tasklists = 'Available calendars: \n'
                for calendar in calendar_list:
                    if calendar['summary'] not in self.calendar_rooms.get(room.room_id, []):
                        tasklists += f" - {calendar['summary']}\n"
                await bot.send_text(room, tasklists)
                return

        elif len(args) == 3:
            if args[1] == 'add':
                bot.must_be_admin(room, event)

                tasklistid = args[2]
                self.logger.info(f'Adding calendar {tasklistid} to room id {room.room_id}')

                if self.calendar_rooms.get(room.room_id):
                    if tasklistid not in self.calendar_rooms[room.room_id]:
                        self.calendar_rooms[room.room_id].append(tasklistid)
                    else:
                        await bot.send_text(room, 'This google calendar already added in this room!')
                        return
                else:
                    self.calendar_rooms[room.room_id] = [tasklistid]

                self.logger.info(f'Calendars now for this room {self.calendar_rooms.get(room.room_id)}')

                bot.save_settings()

                await bot.send_text(room, 'Added new google calendar to this room')
                return

    async def matrix_poll(self, bot, pollcount):
        if pollcount % 6 == 0:  # every minute
            pass


class GoogleTasksList:
    def __init__(self, tasklist_id: str, name: str) -> None:
        self.name: str = name
        self.tasklist_id: str = tasklist_id
        self.google_tasks: Dict[str, GoogleTask] = {}
        self.orphans: List[GoogleTask] = []

    def append_task(self, google_task: "GoogleTask"):
        if google_task.parent_id != "":
            if google_task.parent_id in self.google_tasks:
                self.google_tasks[google_task.parent_id].append_subtask(
                    google_task)
            else:
                self.orphans.append(google_task)
        else:
            self.google_tasks[google_task.task_id] = google_task
            for task in self.orphans.copy():
                if task.parent_id == google_task.task_id:
                    google_task.append_subtask(task)
                    self.orphans.remove(task)

    def to_html(self) -> str:
        html: str = ""
        cpt = 1
        for google_task in self.google_tasks.values():
            task_html, cpt = google_task.to_html(cpt)
            html += task_html + "\n"
            cpt += 1

        for orphan in self.orphans:  # add orphans removing parent link
            orphan.parent_id = ""
            task_html, cpt = orphan.to_html(cpt)
            html += task_html + "\n"
            cpt += 1

        return html

    def get_open_tasks(self) -> int:
        nb_open_tasks = len(self.orphans)
        for google_task in self.google_tasks.values():
            nb_open_tasks += google_task.get_nb_open_tasks()
        return nb_open_tasks


class GoogleTaskLink():
    def __init__(self, link: str, description: str, link_type: str) -> None:
        self.link: str = link
        self.decription: str = description
        self.link_type: str = link_type
        if link_type == "email":
            self.link_type = "✉️ "


class GoogleTask():
    def __init__(self, tasklist_id: str, tasklist: GoogleTasksList,  task_id: str, title: str, notes: str = "", parent_id: str = "", due: str = "", completed: bool = False, favorite: bool = False) -> None:
        self.tasklist_id: str = tasklist_id
        self.task_id: str = task_id
        self.title: str = title
        self.notes: str = notes
        self.parent_id: str = parent_id
        self.tasklist: GoogleTasksList = tasklist
        self.due: Optional[datetime] = None
        self.completed: bool = completed
        self.favorite: bool = favorite
        self.links: List[GoogleTaskLink] = []
        self.sub_tasks: List[GoogleTask] = []
        if due != "":
            self.due = datetime.strptime(due, "%Y-%m-%dT%H:%M:%S.%fZ")

    def add_link(self, link: str, description: str, link_type: str):
        self.links.append(GoogleTaskLink(
            link=link, description=description, link_type=link_type))

    def get_due_date(self) -> str:
        due_date: str = ""
        if self.due is not None:
            due_date = self.due.strftime("%Y-%m-%d")

        return due_date

    def append_subtask(self, google_task: "GoogleTask"):
        self.sub_tasks.append(google_task)

    def to_html(self, cpt: int, deep: int = 0) -> Tuple[str, int]:
        html: str = ""
        favorite = "" if not self.favorite else "⭐ "
        subtitle = ""
        urgentness = ""
        mrkdwn = " " if not self.completed else "~"

        if self.completed:
            urgentness = "⚫ "
        else:
            urgentness = "⚪ "
            if self.due is not None:
                if not self.completed:
                    if self.due < datetime.now():
                        urgentness = "🔴 "
                    elif self.due.day <= datetime.now().day + 3 and self.due.month == datetime.now().month and self.due.year == datetime.now().year:
                        urgentness = "🟠 "
                subtitle = "️ _due " + get_timeago(self.due) + "_"

        for link in self.links:
            subtitle += f"\n         _{link.link_type} <{link.link}|{link.decription}>_"

        html = str(cpt) + " - " + (" " * deep + "└ " if self.parent_id !=
                                   "" else "") + urgentness + favorite + f"{mrkdwn}{self.title}{mrkdwn}" + subtitle

        for sub_task in self.sub_tasks:
            cpt += 1
            sub_html, cpt = sub_task.to_html(cpt, deep+1)
            html += "\n" + sub_html

        return html, cpt

    def get_nb_open_tasks(self) -> int:
        nb_open_tasks = 1
        for task in self.sub_tasks:
            nb_open_tasks += task.get_nb_open_tasks()

        return nb_open_tasks


def _get_credentials():
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def get_tasklists(completed=False) -> List[GoogleTasksList]:
    tasks_lists: List[GoogleTasksList] = []

    creds = _get_credentials()

    try:
        service = build('tasks', 'v1', credentials=creds)
        results = service.tasklists().list(maxResults=10).execute()
        items = results.get('items', [])

        if items:
            for item in items:
                gtasks_json = service.tasks().list(
                    tasklist=item["id"], showCompleted=completed, showDeleted=False, showHidden=False).execute()
                google_task_list: GoogleTasksList = GoogleTasksList(item["id"],
                                                                    item["title"])
                tasks_lists.append(google_task_list)
                items = gtasks_json["items"]
                for task in items:
                    task["dt"] = "99991231"
                    if "due" in task:
                        task["dt"] = datetime.strptime(
                            task["due"], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y%m%d")

                items = sorted(items, key=lambda x: x["dt"])
                for task in items:
                    if completed ^ (task.get("status", "") != "completed"):
                        gtask = GoogleTask(
                            task_id=task["id"],
                            title=task.get("title", ""),
                            notes=task.get("notes", ""),
                            parent_id=task.get("parent", ""),
                            due=task.get("due", ""),
                            completed=completed,
                            tasklist=google_task_list,
                            tasklist_id=item["id"])

                        additional_link = _get_additional_link_from_notes(task)
                        if additional_link is not None:
                            gtask.add_link(
                                additional_link["link"], additional_link["description"], additional_link["type"])
                        for link in task.get("links", []):
                            gtask.add_link(link.get("link", ""), link.get(
                                "description", ""), link.get("type", ""))
                        google_task_list.append_task(gtask)

    except HttpError as err:
        print(err)
    finally:
        return tasks_lists


def toggle_task_completion(tasklist_id: str, task_id: str) -> bool:
    creds = _get_credentials()
    completed: bool = False
    try:
        service = build('tasks', 'v1', credentials=creds)
        task = service.tasks().get(tasklist=tasklist_id, task=task_id).execute()
        status = task["status"]
        if status != "completed":
            task["status"] = "completed"
            completed = True
        else:
            task["status"] = "needsAction"
            completed = False

        service.tasks().update(tasklist=tasklist_id, task=task_id, body=task).execute()

    except HttpError as err:
        print(err)

    return completed


def toggle_task_favorite(tasklist_id: str, task_id: str) -> bool:
    # TODO when API will be updated
    return False


def _get_additional_link_from_notes(task: dict) -> Optional[Dict[str, str]]:
    # if last line of notes starts with :slack: <http
    # we use it as link in the display (hack because gtask API does not allow creating task with links)
    notes = task.get("notes", "")
    additional_link: Optional[Dict[str, str]] = None
    last_note_line = notes.split("\n")[-1]
    if last_note_line.startswith(":slack: <http"):
        additional_link = {}
        split = last_note_line.split("<")[1].split("|")
        additional_link["description"] = split[1][:-1]
        additional_link["link"] = split[0]
        additional_link["type"] = "slack"

    return additional_link


def get_task(tasklist_id: str, task_id: str) -> Optional[GoogleTask]:
    gtask: Optional[GoogleTask] = None
    creds = _get_credentials()

    try:
        service = build('tasks', 'v1', credentials=creds)
        tasklist: Optional[GoogleTasksList] = get_tasklist(tasklist_id)
        if tasklist is not None:
            task = service.tasks().get(tasklist=tasklist_id, task=task_id).execute()
            gtask = GoogleTask(
                task_id=task["id"], title=task.get("title", ""), notes=task.get("notes", ""), parent_id=task.get("parent", ""), due=task.get("due", ""), completed=(task["status"] == "completed"), tasklist_id=task["id"], tasklist=tasklist)

            additional_link = _get_additional_link_from_notes(task)
            if additional_link is not None:
                gtask.add_link(
                    additional_link["link"], additional_link["description"], additional_link["type"])
            for link in task.get("links", []):
                gtask.add_link(link.get("link", ""), link.get(
                    "description", ""), link.get("type", ""))
    except HttpError as err:
        print(err)

    return gtask


def update_task(old_tasklist_id: str, new_tasklist_id: str, task_id: str, task_title: str, task_description: str, task_duedate: str) -> bool:
    creds = _get_credentials()
    completed: bool = False
    try:
        service = build('tasks', 'v1', credentials=creds)
        task = service.tasks().get(tasklist=old_tasklist_id, task=task_id).execute()
        task["title"] = task_title
        task["notes"] = task_description
        if task_duedate != "":
            # time is discarded in the google task api
            task["due"] = f"{task_duedate}T00:00:00.000Z"

        if old_tasklist_id == new_tasklist_id:
            service.tasks().update(tasklist=new_tasklist_id, task=task_id, body=task).execute()
        else:
            delete_task(tasklist_id=old_tasklist_id, task_id=task_id)
            del task["id"]
            service.tasks().insert(tasklist=new_tasklist_id, body=task).execute()

        if task["status"] == "completed":
            completed = True

    except HttpError as err:
        print(err)

    return completed


def create_task(task_title: str, task_description: str, task_duedate: Optional[str], tasklist_id: str = "", task_links: List[str] = []):
    creds = _get_credentials()
    try:
        service = build('tasks', 'v1', credentials=creds)
        task: dict = {}
        task["title"] = task_title
        task["notes"] = task_description
        if task_duedate is not None and task_duedate != "":
            # time is discarded in the google task api
            task["due"] = f"{task_duedate}T00:00:00.000Z"

        if len(task_links) > 0:
            task["links"] = task_links

        service.tasks().insert(tasklist=tasklist_id, body=task).execute()

    except HttpError as err:
        print(err)


def delete_task(tasklist_id: str, task_id: str) -> bool:
    creds = _get_credentials()
    completed: bool = False
    try:
        service = build('tasks', 'v1', credentials=creds)
        task = service.tasks().get(tasklist=tasklist_id, task=task_id).execute()
        if task["status"] == "completed":
            completed = True

        service.tasks().delete(tasklist=tasklist_id, task=task_id).execute()

    except HttpError as err:
        print(err)

    return completed


def get_tasklist_by_title(tasklist_title: str) -> Optional[GoogleTasksList]:
    creds = _get_credentials()
    google_tasklist: Optional[GoogleTasksList] = None
    try:
        service = build('tasks', 'v1', credentials=creds)
        items = service.tasklists().list().execute()
        for item in items["items"]:
            if item.get("title", "") == tasklist_title:
                google_tasklist = GoogleTasksList(item["id"], item["title"])
                break

    except HttpError as err:
        print(err)

    return google_tasklist


def get_tasklist(tasklist_id: str) -> Optional[GoogleTasksList]:
    tasklist: Optional[GoogleTasksList] = None
    creds = _get_credentials()
    try:
        service = build('tasks', 'v1', credentials=creds)
        tasklist_json = service.tasklists().get(tasklist=tasklist_id).execute()
        tasklist = GoogleTasksList(
            tasklist_json["id"], tasklist_json.get("title", ""))

    except HttpError as err:
        print(err)

    return tasklist


def create_tasklist(tasklistname: str):
    creds = _get_credentials()
    try:
        service = build('tasks', 'v1', credentials=creds)
        tasklist: dict = {}
        tasklist["title"] = tasklistname
        service.tasklists().insert(body=tasklist).execute()

    except HttpError as err:
        print(err)


def get_timeago(dt: datetime) -> str:
    str_timeago: str = ""
    if dt.day == datetime.now().day and dt.month == datetime.now().month and dt.year == datetime.now().year:
        str_timeago = "today"
    elif dt.day == datetime.now().day+1 and dt.month == datetime.now().month and dt.year == datetime.now().year:
        str_timeago = "tomorrow"
    elif dt.day == datetime.now().day-1 and dt.month == datetime.now().month and dt.year == datetime.now().year:
        str_timeago = "yesterday"
    else:
        str_timeago = timeago.format(dt)

    return str_timeago


if __name__ == "__main__":
    for task_list in get_tasklists():
        print(task_list.to_html())

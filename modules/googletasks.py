from __future__ import print_function
from typing import Dict, List, Optional, Tuple
import os.path
from datetime import datetime
import timeago
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from nio import MatrixRoom
from modules.common.module import BotModule
#
# Google task notifications
#


SCOPES = ['https://www.googleapis.com/auth/tasks']


class MatrixModule(BotModule):

    def __init__(self, name):
        super().__init__(name)
        self.service_name = 'googletasks'
        self.credentials_file = "credentials.json"
        self.tasklists_rooms = {}
        self.enabled = True
        self.service = None
        self.poll_interval_min = 1

        # contains last given tasks for a room in order to easily pick them by number instead of their long ids
        self.tasks_per_room: Dict[str, List["GoogleTask"]] = {}

    def matrix_start(self, bot):
        super().matrix_start(bot)

        if not os.path.exists(self.credentials_file) or os.path.getsize(self.credentials_file) == 0:
            return  # No-op if not set up

        creds = _get_credentials()
        self.service = build('tasks', 'v1', credentials=creds)

        try:
            results = self.service.tasklists().list().execute()
            tasklists = results.get('items', [])
            if tasklists:
                for tasklist in tasklists:
                    self.logger.info("%s - %s", {tasklist['title']}, {tasklist['id']})
        except Exception:
            self.logger.error('Getting tasklist list failed!')

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
            await bot.send_text(room, 'Google tasklist not set up for this bot.')
            return
        args = event.body.split()
        tasklist_names = self.tasklists_rooms.get(room.room_id) or []

        if len(args) == 2:
            if args[1] == 'today':
                await self.cmd_list_today(bot, room, tasklist_names, display_tasklist_if_empty=True)
            elif args[1] == 'list':
                await bot.send_text(room, 'Tasklists in this room: ' + str(self.tasklists_rooms.get(room.room_id)))
            elif args[1] == 'listavailable':
                await self.cmd_listavailable(bot, room)
        elif len(args) >= 3:
            if args[1] == 'add':
                await self.cmd_add_tasklist_to_room(bot, room, event, ' '.join(args[2:]))
            elif args[1] == 'del':
                await self.cmd_del_tasklist_from_room(bot, room, event, ' '.join(args[2:]))
            elif args[1] == 'show':
                await self.cmd_show_task_by_index(bot, room, args[2])
            else:
                await bot.send_text(room, 'Unknown command')
        else:
            await self.cmd_list(bot, room, tasklist_names, None, display_tasklist_if_empty=True)

    async def cmd_list(self, bot, room, tasklist_names: List[str], until_date: Optional[datetime], display_tasklist_if_empty: bool):
        html = ''
        text = ''
        self.logger.info('cmd_list, len of tasklist_names (%d)', len(tasklist_names))
        for tasklist_name in tasklist_names:
            cpt = 1
            self.logger.info('Listing tasks in tasklist "%s"', tasklist_name)
            tasklist_html = f'<h4>{tasklist_name}</h4>\n'
            tasklist_text = f'{tasklist_name}\n'
            self.tasks_per_room[room.room_id] = []
            for task in self._get_tasks_for_tasklist_and_date(tasklist_name, until_date):
                self.tasks_per_room[room.room_id].append(task)
                task_html, task_text = task.get_html_and_text_summary()
                tasklist_html += f'{cpt} - {task_html}'
                tasklist_text += f'{cpt} - {task_text}'

                cpt += 1

            if cpt == 1:
                tasklist_html += 'Nothing to do! 😙'
                tasklist_text += 'Nothing to do! 😙'
            if cpt > 1 or display_tasklist_if_empty:
                html += tasklist_html
                text += tasklist_text

        if html != '':
            await bot.send_html(room, html, text)

    async def cmd_list_today(self, bot, room, tasklist_names: List[str], display_tasklist_if_empty: bool):
        await self.cmd_list(bot, room, tasklist_names, datetime.now(), display_tasklist_if_empty)

    async def cmd_show_task_by_index(self, bot, room: MatrixRoom, index: str):
        if index.isdigit():
            if room.room_id in self.tasks_per_room:
                if int(index) <= len(self.tasks_per_room[room.room_id]):
                    html, text = self.tasks_per_room[room.room_id][int(index)-1].get_html_and_text_full_description()
                    await bot.send_html(room, html, text)
                else:
                    await bot.send_text(room, f'Invalid index. Expected index from 1 to {len(self.tasks_per_room[room.room_id])}')
            else:
                await bot.send_text(room, 'No task list has been displayed in this room. First display task lists.')
        else:
            await bot.send_text(room, f'\'{index}\' is not a valid index.\nUsage !googletasks show <idx (integer)>')

    async def cmd_del_tasklist_from_room(self, bot, room, event, tasklist_name):
        bot.must_be_admin(room, event)

        self.logger.info('Removing calendar %s from room id %s', tasklist_name, room.room_id)

        if self.tasklists_rooms.get(room.room_id) and tasklist_name in self.tasklists_rooms[room.room_id]:
            self.tasklists_rooms[room.room_id].remove(tasklist_name)

        self.logger.info('Tasklists now for this room %s', self.tasklists_rooms.get(room.room_id))

        bot.save_settings()

        await bot.send_text(room, 'Removed google tasklist from this room')

    async def cmd_add_tasklist_to_room(self, bot, room, event, tasklist_name: str):
        bot.must_be_admin(room, event)

        self.logger.info('Adding tasklist %s to room id %s', tasklist_name, room.room_id)

        if self.tasklists_rooms.get(room.room_id):
            if tasklist_name not in self.tasklists_rooms[room.room_id]:
                if self._get_tasklist_by_title(tasklist_name) is not None:
                    self.tasklists_rooms[room.room_id].append(tasklist_name)
                else:
                    await bot.send_text(room, 'This google tasklist does not exist!')
                    return
            else:
                await bot.send_text(room, 'This google tasklist already added in this room!')
                return
        else:
            self.tasklists_rooms[room.room_id] = [tasklist_name]

        self.logger.info('Tasklist now for this room %s', self.tasklists_rooms.get(room.room_id))

        bot.save_settings()

        await bot.send_text(room, 'Added new google tasklist to this room')

    async def cmd_listavailable(self, bot, room):
        tasklist_list = self.service.tasklists().list().execute()['items']
        tasklists = 'Available tasklists: \n'
        for tasklist in tasklist_list:
            if tasklist['title'] not in self.tasklists_rooms.get(room.room_id, []):
                tasklists += f" - {tasklist['title']}\n"
        await bot.send_text(room, tasklists)

    def _get_tasks_for_tasklist_and_date(self, tasklist_name: str, until_date: Optional[datetime] = None) -> List["GoogleTask"]:
        """
        return tasks for given task list
        in case of until_date is defined, only tasks having a due date older or same as until_date are returned
        """
        tasks: List[GoogleTask] = []
        tasklist: Optional["GoogleTasksList"] = self._get_tasklist_by_title(tasklist_name)
        if tasklist is not None:
            for task in tasklist.google_tasks.values():
                if until_date is None or (task.due is not None and task.due <= until_date):
                    tasks.append(task)

            for orphan in tasklist.orphans:  # add orphans removing parent link
                if until_date is None or (orphan.due is not None and orphan.due <= until_date):
                    tasks.append(orphan)

        return tasks

    async def matrix_poll(self, bot, pollcount):
        """
        Display a list of tasks due for today and overdue tasks, at 7:30am except during weekends
        """
        if pollcount % 6 == 0:  # every minute
            now = datetime.now()
            if now.weekday() not in [5, 6] and now.hour == 7 and now.minute == 30:
                for room_id in self.tasklists_rooms:
                    self.logger.info('Display digest for room "%s" if any task is due', room_id)
                    room = MatrixRoom(room_id=room_id, own_user_id='')
                    await self.cmd_list_today(bot, room, self.tasklists_rooms.get(room_id, []), display_tasklist_if_empty=False)

    def _get_tasklist_by_title(self, tasklist_title: str, completed: bool = False) -> Optional["GoogleTasksList"]:
        google_tasklist: Optional[GoogleTasksList] = None
        try:
            gtasklist_json = self.service.tasklists().list().execute()
            for item in gtasklist_json["items"]:
                if item.get("title", "") == tasklist_title:
                    google_tasklist = GoogleTasksList(item["id"], item["title"])

                    self._load_tasklist_tasks(google_tasklist, completed)

        except HttpError as err:
            self.logger.error(err)

        return google_tasklist

    def _load_tasklist_tasks(self, google_tasklist: "GoogleTasksList", completed: bool = False):
        """
        Loads tasks for a given tasklist.
        @param completed: by default completed tasks are not loaded except if this param is True
        """
        gtasks_json = self.service.tasks().list(
            tasklist=google_tasklist.tasklist_id, showCompleted=completed, showDeleted=False, showHidden=False).execute()

        tasks_items = gtasks_json["items"]
        for task in tasks_items:
            task["dt"] = "99991231"
            if "due" in task:
                task["dt"] = datetime.strptime(
                    task["due"], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%Y%m%d")

        tasks_items = sorted(tasks_items, key=lambda x: x["dt"])
        for task in tasks_items:
            if completed ^ (task.get("status", "") != "completed"):
                gtask = GoogleTask(
                    task_id=task["id"],
                    title=task.get("title", ""),
                    notes=task.get("notes", ""),
                    parent_id=task.get("parent", ""),
                    due=task.get("due", ""),
                    completed=completed,
                    tasklist=google_tasklist,
                    tasklist_id=google_tasklist.tasklist_id)

                additional_link = _get_additional_link_from_notes(task)
                if additional_link is not None:
                    gtask.add_link(
                        additional_link["link"], additional_link["description"], additional_link["type"])
                for link in task.get("links", []):
                    gtask.add_link(link.get("link", ""), link.get(
                        "description", ""), link.get("type", ""))
                google_tasklist.append_task(gtask)

    def _get_tasklist(self, tasklist_id: str) -> Optional["GoogleTasksList"]:
        tasklist: Optional[GoogleTasksList] = None
        try:
            tasklist_json = self.service.tasklists().get(tasklist=tasklist_id).execute()
            tasklist = GoogleTasksList(
                tasklist_json["id"], tasklist_json.get("title", ""))

        except HttpError as err:
            self.logger.error(err)

        return tasklist

    def _get_task_by_id(self, task_id: str) -> Optional["GoogleTask"]:
        task: Optional[GoogleTask] = None
        try:
            tasklist_json = self.service.tasklists().get(tasklist=tasklist_id).execute()
            task = GoogleTasksList(
                tasklist_json["id"], tasklist_json.get("title", ""))

        except HttpError as err:
            self.logger.error(err)

        return task

    def help(self):
        return 'Google tasklist. Lists 10 next tasks by default. today = list today\'s tasks. (Available commands: today, list, listavailable, add, del)'

    def long_help(self, bot=None, event=None, **kwargs):
        text = self.help() + (
            '\n- "!googletask list": list tasklists associated to current room'
            '\n- "!googletask listavailable": list available tasklists for current room (omit already associated)'
            '\n- "!googletask today": list today tasks for every associated tasklist')
        if bot and event and bot.is_owner(event):
            text += (
                '\n- "!googletask add <tasklist name>": associate the tasklist to the room'
                '\n- "!googletask del <tasklist name>": de-associate the tasklist to the room'
            )
        text += '\n\n Every day at 7:30am the bot will send a digest with due for today tasks and overdue tasks for all rooms having associated tasklists'
        return text


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

    def get_html_and_text_summary(self, deep: int = 0) -> Tuple[str, str]:
        html: str = ""
        text: str = ""
        favorite = "" if not self.favorite else "⭐ "
        html_subtitle = ""
        text_subtitle = ""
        urgentness = ""
        html_title = self.title
        text_title = self.title
        if self.completed:
            html_title = f"<s>{self.title}</s>"
            text_title += " (completed)"
            urgentness = "⚫ "

        else:
            urgentness = "⚪ "
            if self.due is not None:
                if not self.completed:
                    if self.due < datetime.now():
                        urgentness = "🔴 "
                    elif self.due.day <= datetime.now().day + 3 and self.due.month == datetime.now().month and self.due.year == datetime.now().year:
                        urgentness = "🟠 "
                html_subtitle = "️ <i>due " + get_timeago(self.due) + "</i>"
                text_subtitle = "️ (due " + get_timeago(self.due) + ")"

        for link in self.links:
            html_subtitle += f'\n<br/>         <i>{link.link_type} <a href="{link.link}">{link.decription}</a></i>'
            text_subtitle += f'\n         {link.link_type} {link.decription} {link.link}'

        html = (" " * deep + "└ " if self.parent_id !=
                "" else "") + urgentness + favorite + " " + html_title + html_subtitle + '<br/>'

        text = (" " * deep + "└ " if self.parent_id !=
                "" else "") + urgentness + favorite + " " + text_title + text_subtitle + '\n'

        for sub_task in self.sub_tasks:
            sub_html, sub_text = sub_task.get_html_and_text_summary(deep+1)
            html += "\n<br/>" + sub_html
            text += "\n" + sub_text

        return html, text

    def get_html_and_text_full_description(self) -> Tuple[str, str]:
        html: str = ''
        text: str = ''
        favorite = '' if not self.favorite else '⭐'
        html_subtitle = ''
        text_subtitle = ''
        urgentness = ''
        html_title = self.title
        text_title = self.title
        if self.completed:
            html_title = f'<s>{self.title}</s>'
            text_title += ' (completed)'
            urgentness = '⚫ '

        else:
            urgentness = '⚪ '
            if self.due is not None:
                if not self.completed:
                    if self.due < datetime.now():
                        urgentness = '🔴 '
                    elif self.due.day <= datetime.now().day + 3 and self.due.month == datetime.now().month and self.due.year == datetime.now().year:
                        urgentness = '🟠 '
                html_subtitle = '📅 due ' + get_timeago(self.due) + '<br/>'
                text_subtitle = '📅 due ' + get_timeago(self.due)

        for link in self.links:
            html_subtitle += f'         <i>{link.link_type} <a href="{link.link}">{link.decription}</a></i><br/>'
            text_subtitle += f'\n         {link.link_type} {link.decription} {link.link}'

        html = f'<hr/><h3>{urgentness} {favorite} {html_title}</h3>{html_subtitle}------------<br/>{self.notes}'

        text = f'{urgentness} {favorite} {text_title}\n{text_subtitle}\n{self.notes}'

        """
        for sub_task in self.sub_tasks:
            sub_html, sub_text = sub_task.get_html_and_text_summary()
            html += '\n<br/>' + sub_html
            text += '\n' + sub_text
        """

        return html, text

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

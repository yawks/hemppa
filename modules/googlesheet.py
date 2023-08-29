from __future__ import print_function
from typing import Dict, List, Optional, Tuple
import os.path
import gspread
from datetime import datetime #do not remove, maybe used in exec()
import pandas as pd
from gspread import Client
from oauth2client.service_account import ServiceAccountCredentials
from nio import MatrixRoom
from modules.common.module import BotModule
#
# Google sheet reader
#

SCOPE = 'https://spreadsheets.google.com/feeds'


class MatrixModule(BotModule):

    def __init__(self, name):
        super().__init__(name)
        self.service_name = 'googlesheet'
        self.credentials_file = "client_secret.json"

        self.sheets_rooms: Dict[str, Dict[str, str]] = {}
        self.command_aliases: Dict[str, Dict[str, Tuple[str, str]]] = {}
        self.enabled = True
        self.client = None
        self.poll_interval_min = 1

    def matrix_start(self, bot):
        super().matrix_start(bot)

        if not os.path.exists(self.credentials_file) or os.path.getsize(self.credentials_file) == 0:
            return  # No-op if not set up

        self.client = self._get_spread_sheet_client()

    def get_settings(self):
        data = super().get_settings()
        data['sheets_rooms'] = self.sheets_rooms
        data['sheets_command_aliases'] = self.command_aliases
        return data

    def set_settings(self, data):
        super().set_settings(data)
        if data.get('sheets_rooms'):
            self.sheets_rooms = data['sheets_rooms']
        if data.get('sheets_command_aliases'):
            self.command_aliases = data['sheets_command_aliases']

    async def matrix_message(self, bot, room: MatrixRoom, event):
        if not self.client:
            await bot.send_text(room, 'Google sheet reader not set up for this bot.')
            return
        args = _get_args(event.body)
        if len(args) > 1:
            match args[1]:
                case 'help':
                    await bot.send_text(room, self.long_help(bot))
                case 'list':
                    #!googlesheet list
                    await self.cmd_list_room_sheets(bot, room)
                case 'add':
                    #!googlesheet add <doc id> "<doc name>"
                    await self.cmd_add_sheet_to_room(bot, room, event, args[2], ' '.join(args[3:]))
                case 'del':
                    #!googlesheet del "<doc name>"
                    # also delete every alias attached to this doc
                    await self.cmd_del_sheet_from_room(bot, room, event, args[2])
                case 'alias':
                    #!googlesheet alias "<doc name>" "<alias name>" "<panda query>"
                    # panda query (can be multiline), the dataframe variable is called df
                    #   ie: df.groupby("col").sum()
                    await self.cmd_command_alias(bot, room, event, args)
                case 'aliases':
                    #!googlesheet alias
                    # list aliases for current room
                    await self.cmd_list_room_aliases(bot, room)
                case _:
                    #!googlesheet <alias name>
                    # execute alias command
                    await self.cmd_execute_alias(bot, room, args)
        else:
            await bot.send_text(room, 'You must specify a command.\n' + self.long_help(bot, event))

    async def cmd_execute_alias(self, bot, room: MatrixRoom, args: List[str]):
        alias_name = ' '.join(args[1:])
        if room.room_id not in self.command_aliases or alias_name not in self.command_aliases[room.room_id]:
            await bot.send_text(room, f'Uknown alias {alias_name}')
            await self.cmd_list_room_aliases(bot, room)
        else:
            sheet_id, query = self.command_aliases[room.room_id][alias_name]
            sheet = self.client.open_by_url(f'https://docs.google.com/spreadsheets/d/{sheet_id}')

            worksheet = sheet.worksheet('CA par année')
            df = pd.DataFrame(worksheet.get_all_records())
            _locals = locals()
            try:
                exec(query, globals(), _locals)
                if _locals.get('result', None) is not None:
                    text = str(_locals['result'])
                    html = f'<pre>{text}</pre>'
                    await bot.send_html(room, html, text)
                else:
                    await bot.send_text(room, 'Your query should set a variable named "result".')
            except Exception as e:
                await bot.send_text(room, f'Error executing query:\n{query}\n-------------\nException:\n {repr(e)}')

    async def cmd_list_room_aliases(self, bot, room: MatrixRoom):
        self.logger.info('List aliases for room id %s', room.room_id)
        text = ''
        if room.room_id in self.command_aliases and len(self.command_aliases[room.room_id]) > 0:
            text = 'Available aliases:'
            for alias in self.command_aliases[room.room_id]:
                text += f'\n - {alias} : {self.command_aliases[room.room_id][alias][1]}'
        else:
            text = 'This room has no available alias. Add one first.'

        await bot.send_text(room, text)

    async def cmd_list_room_sheets(self, bot, room: MatrixRoom):
        self.logger.info('List sheets for room id %s', room.room_id)
        text = ''
        if room.room_id in self.sheets_rooms and len(self.sheets_rooms[room.room_id]) > 0:
            text = 'Available sheets:'
            for s_id, s_name in self.sheets_rooms[room.room_id].items():
                text += f'\n - {s_name} : {s_id}'
        else:
            text = 'This room has no available sheet. Add one first.'

        await bot.send_text(room, text)

    async def cmd_command_alias(self, bot, room: MatrixRoom, event, args: List[str]):
        bot.must_be_admin(room, event)
        self.logger.info('Alias for room id %s => %s', room.room_id, ' '.join(args))

        if len(args) > 2:
            action = args[2]
            if action == "add":
                await self._cmd_add_alias(bot, room, args)
            elif action == "del":
                await self._cmd_del_alias(bot, room, event, args)
            else:
                await bot.send_text(room, f'Unknown action: {action}.\nAvailable actions: add, del')

    async def _cmd_del_alias(self, bot, room, event, args):
        if len(args) > 3:
            alias_name = args[3]
            if room.room_id not in self.command_aliases or alias_name not in self.command_aliases[room.room_id]:
                await bot.send_text(room, f'Uknown alias {alias_name}')
                await self.cmd_list_room_aliases(bot, room)
            else:
                del self.command_aliases[room.room_id][alias_name]
                bot.save_settings()
                await bot.send_text(room, f'Removed alias "{alias_name}"')
        else:
            await bot.send_text(room, 'Not enough arguments\n!googlesheet alias del "<alias name>"')

    async def _cmd_add_alias(self, bot, room, args):
        if len(args) > 4:
            sheet_name = args[3]
            alias_name = args[4]
            command = ' '.join(args[5:])
            sheet_id: str | None = self._get_sheet_id_by_name(room, sheet_name)

            if sheet_id is not None:
                if room.room_id not in self.command_aliases:
                    self.command_aliases[room.room_id] = {}
                if alias_name not in self.command_aliases[room.room_id]:
                    self.command_aliases[room.room_id][alias_name] = (sheet_id, command)
                    bot.save_settings()
                    await bot.send_text(room, f'Added alias \'{alias_name}\' for sheet \'{sheet_name}\' to this room')

                else:
                    await bot.send_text(room, f'This alias \'{alias_name}\' already exists')

            else:
                await bot.send_text(room, f'Google sheet \'{sheet_name}\' does not exist!')
        else:
            await bot.send_text(room, 'Not enough arguments\n!googlesheet alias add "<doc name>" "<alias name>" panda query')

    async def cmd_del_sheet_from_room(self, bot, room: MatrixRoom, event, sheet_name: str):
        bot.must_be_admin(room, event)
        self.logger.info('Delete sheet "%s" from room id %s', sheet_name, room.room_id)

        sheet_id = self._get_sheet_id_by_name(room, sheet_name)
        if sheet_id is not None:
            del self.sheets_rooms[room.room_id][sheet_id]
            # also remove aliases of this doc
            for alias in list(self.command_aliases[room.room_id]):
                doc_id, _ = self.command_aliases[room.room_id][alias]
                if doc_id == sheet_id:
                    del self.command_aliases[room.room_id][alias]

            self.logger.info('Sheet now deleted for this room %s', self.sheets_rooms.get(room.room_id))
            bot.save_settings()
            await bot.send_text(room, 'Removed google sheet from this room')
        else:
            await bot.send_text(room, f'Sheet {sheet_name} does not exist in this room')

    async def cmd_add_sheet_to_room(self, bot, room: MatrixRoom, event, sheet_id: str, sheet_name: str):
        bot.must_be_admin(room, event)

        self.logger.info('Adding sheet %s \'%s\' to room id %s', sheet_id, sheet_name, room.room_id)

        if self.sheets_rooms.get(room.room_id):
            if sheet_id not in self.sheets_rooms[room.room_id]:
                if self._get_sheet_by_id(sheet_id) is not None:
                    self.sheets_rooms[room.room_id][sheet_id] = sheet_name
                else:
                    await bot.send_text(room, 'This google sheet does not exist!')
                    return
            else:
                await bot.send_text(room, 'This google sheet already added in this room!')
                return
        else:
            self.sheets_rooms[room.room_id] = {}
            self.sheets_rooms[room.room_id][sheet_id] = sheet_name

        self.logger.info('Sheet now for this room %s', self.sheets_rooms.get(room.room_id))

        bot.save_settings()

        await bot.send_text(room, 'Added new google sheet to this room')

    def _get_sheet_by_id(self, sheet_id) -> gspread.Spreadsheet | None:
        sheet: gspread.Spreadsheet | None = None
        try:
            sheet = self.client.open_by_url('https://docs.google.com/spreadsheets/d/1QwjDkh_uvf203qjVxMIhQfvlNlLAzZ5IWO8BQEVHOFQ')
        except gspread.SpreadsheetNotFound as _:
            self.logger.error('Sheet \'%s\' not found', sheet_id)

        return sheet

    def _get_spread_sheet_client(self) -> Optional[Client]:
        client: Optional[Client] = None
        # use creds to create a client to interact with the Google Drive API
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_file, SCOPE)
            client = gspread.authorize(creds)
        except Exception as exception:
            self.logger.error("Error loading sheet %s", str(exception))

        return client

    def _get_sheet_id_by_name(self, room: MatrixRoom, sheet_name) -> str | None:
        sheet_id: str | None = None
        if self.sheets_rooms.get(room.room_id):
            for s_id, s_name in self.sheets_rooms[room.room_id].items():
                if s_name == sheet_name:
                    sheet_id = s_id
                    break

        return sheet_id

    def help(self):
        return 'Google sheet. Can read google spreadsheet tables and display them in conversation. Panda module can be used to transform data.'

    def long_help(self, bot=None, event=None, **kwargs):
        text = self.help() + (
            '\n- "!googlesheet list": list sheets associated to current room'
            '\n- "!googlesheet aliases": list availabled aliases of the current room'
            '\n- "!googlesheet <alias name>": execute the panda query of the given alias')
        if bot and event and bot.is_owner(event):
            text += (
                '\n- "!googlesheet add <doc id> <doc name>": add a doc. The <doc name> will used for following commands.'
                '\n- "!googlesheet alias add <doc name> <alias name> <panda query>": alias a panda query for the given <doc name>.'
                '\n\t\teg: !googlesheet alias "my sheet" "get_table" result=df.groupby(\'sheet name\').sum().filter([\'Col 1\', \'Col 2\', \'Col 3\']).transpose()'
                '\n\t\t_df_ is the variable name for panda object. Multiple lines can be set for the alias.'
                '\n\t\tA variable named "result" must be set'
                '\n\t\tWarning, this string should not contain double quotes'
                '\n- "!googlesheet alias del <alias name> ": delete alias <alias name>.'
            )
        return text


def _get_args(command: str) -> List[str]:
    args: List[str] = []
    quote_open = False
    current_arg = ""
    for token in command.strip().split(' '):
        if quote_open:
            current_arg += ' ' + token
            if token[-1] == '"':
                args.append(current_arg[1:-1])  # remove quotes
                current_arg = ""
                quote_open = False
        elif token[0] == '"':
            quote_open = True
            current_arg = token
        else:
            args.append(token)

    return args

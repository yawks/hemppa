import os
import deepl
from modules.common.module import BotModule

USAGE = 'Usage: !translate <target_lang> <my message to translate>\nEg: !translate FR Hello, how are you?"'

class MatrixModule(BotModule):
    def __init__(self, name):
        super().__init__(name)
        self.enabled = True
        self.service_name = 'translate'

    async def matrix_message(self, bot, room, event):
        args = event.body.split()
        if len(args) < 3:
            await bot.send_text(USAGE)
        else:
            target_lang = args[1]
            translator = deepl.Translator(os.getenv('DEEPL_KEY', ''), server_url='https://api-free.deepl.com')
            result = translator.translate_text(' '.join(args[2:]), target_lang=target_lang)
            if isinstance(result, list):
                await bot.send_text(room, result[0].text)
            else:
                await bot.send_text(room, result.text)

    def help(self):
        return 'Translate what user has said\n' + USAGE

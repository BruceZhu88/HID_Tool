import json
import base64

class Settings():

    def __init__(self, path):
        self.path = path
        self.settings = {}

    def loadSettings(self):
        try:
            with open(self.path, 'r', encoding = 'utf-8') as js:
                self.settings = json.load(js)
        except FileNotFoundError:
            print('Error: can\'t find file or read data from local settings')

    def saveSettings(self):
        with open(self.path, mode='w', encoding='utf-8') as js:
            json.dump(self.settings, js, indent = 4)

    def get(self, key, default = '', crypt = False):
        value = self.settings.get(key, default)
        if crypt:
            return base64.b64decode(value.encode(encoding = "utf-8")).decode()
        else:
            return value

    def set(self, key, value, crypt = False):
        if crypt:
            value = base64.b64encode(value.encode(encoding = "utf-8")).decode()
        self.settings[key] = value

import json
import wx
import os


def get_config_file_path():
    return wx.StandardPaths.Get().GetUserConfigDir() + "/esp-flasher-gui.json"


class FlashConfig:
    def __init__(self):
        self.baud = 921600
        self.port = None
        self.firmware_path = None
        self.mode = 'dio'
        self.erase_flash = 'No'

    @classmethod
    def load(cls):
        file_path = get_config_file_path()
        print(file_path)
        conf = cls()
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                data = json.load(f)
            conf.port = data['port']
            conf.baud = data['baud']
            conf.mode = data['mode']
            conf.erase_flash = data['erase']
        return conf

    def save(self):
        file_path = get_config_file_path()
        print(file_path)
        date = {
            'baud': self.baud,
            'port': self.port,
            'mode': self.mode,
            'erase': self.erase_flash
        }
        with open(file_path, 'w') as f:
            json.dump(date, f)


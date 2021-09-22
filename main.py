import wx
import esptool
import threading
import serial
import sys
import os
from serial.tools import list_ports

DEVNULL = open(os.devnull, 'w')

__version__ = "0.0.1"

main_port = None
mac = None


class Espflasher(Exception):
    pass


def detect_chip(port):
    try:
        chip = esptool.ESPLoader.detect_chip(port)
    except esptool.FatalError as err:
        raise Espflasher("ESP Chip Auto-Detection failed: {}".format(err))

    try:
        chip.connect()
    except esptool.FatalError as err:
        raise Espflasher("Error connecting to ESP: {}".format(err))

    return chip


def read_chip_property(func, *args, **kwargs):
    try:
        return prevent_print(func, *args, **kwargs)
    except esptool.FatalError as err:
        raise Espflasher("Reading chip details failed: {}".format(err))


def prevent_print(func, *args, **kwargs):
    orig_sys_stdout = sys.stdout
    sys.stdout = DEVNULL
    try:
        return func(*args, **kwargs)
    except serial.SerialException as err:
        raise Espflasher("Serial port closed: {}".format(err))
    finally:
        sys.stdout = orig_sys_stdout
        pass


def esptool_read_mac(port):
    chip = detect_chip(port)
    mac_address = (':'.join('{:02X}'.format(x) for x in read_chip_property(chip.read_mac)))
    print(f'MAC - {mac_address}')
    return mac_address


class EspToolThread(threading.Thread):
    def __init__(self, parent, config, txt_ctrl):
        threading.Thread.__init__(self)
        self.txt_ctrl = txt_ctrl
        # to exit when main thread exits
        self.daemon = True
        # self._parent = parent
        self._config = config
        self.mac = None

    def run(self):
        try:
            # self.read_mac()
            argv = ["--port", self._config.port]
            argv.extend(['--baud', str(self._config.baud)])
            argv.extend(['--after', 'hard_reset', 'write_flash'])
            argv.extend(["--flash_size", "detect",
                         "--flash_mode", self._config.mode,
                         "0x00000", self._config.firmware_path])

            if self._config.erase_flash:
                argv.append('--erase-all')
            print(argv)
            print("Command: esptool.py %s\n" % " ".join(argv))

            esptool.main(argv)

        except Exception as e:
            print("Unexpected error: {}".format(e))
            raise e

    def read_mac(self):
        # read mac and update to UI
        self.mac = esptool_read_mac(self._config.port)
        wx.CallAfter(self.txt_ctrl.SetValue, self.mac)


class FlashConfig:
    def __init__(self):
        self.baud = 921600
        self.port = None
        self.firmware_path = None
        self.mode = 'dio'
        self.erase_flash = False

    @classmethod
    def load(cls):
        conf = cls()
        return conf


class MyPanel(wx.Panel):
    def __init__(self, parent):
        super(MyPanel, self).__init__(parent)

        self._config = FlashConfig.load()

        # labels
        port_label = wx.StaticText(self, label='Serial Port')
        file_label = wx.StaticText(self, label='Firmware file')
        # baud_label = wx.StaticText(self, label = 'Baud rate')
        erase_label = wx.StaticText(self, label='Erase Flash')

        hbox = wx.BoxSizer(wx.HORIZONTAL)

        flex_grid = wx.FlexGridSizer(6, 2, 10, 10)

        self.choice = wx.Choice(self, choices=self._get_serial_ports())
        self.choice.Bind(wx.EVT_CHOICE, self.on_select_port)
        # self._select_configured_port()

        reload_button = wx.Button(self, label="Reload")
        reload_button.Bind(wx.EVT_BUTTON, self.on_reload)
        reload_button.SetToolTip("Reload serial device list")

        serial_boxsizer = wx.BoxSizer(wx.HORIZONTAL)
        serial_boxsizer.Add(self.choice, 1, wx.EXPAND)
        serial_boxsizer.Add(reload_button, flag=wx.LEFT, border=10)

        file_picker = wx.FilePickerCtrl(self, style=wx.FLP_USE_TEXTCTRL)
        file_picker.Bind(wx.EVT_FILEPICKER_CHANGED, self.on_pick_file)

        erase_no_button = wx.RadioButton(self, label='No', style=wx.RB_GROUP)
        erase_yes_button = wx.RadioButton(self, label='Yes')
        self.Bind(wx.EVT_RADIOBUTTON, self.on_erase_change)

        erase_boxsizer = wx.BoxSizer(wx.HORIZONTAL)
        erase_boxsizer.Add(erase_no_button)
        erase_boxsizer.Add(erase_yes_button)

        upload_button = wx.Button(self, label="Upload Firmware")
        upload_button.Bind(wx.EVT_BUTTON, self.on_upload)

        read_mac_button = wx.Button(self, label="Read MAC")
        read_mac_button.Bind(wx.EVT_BUTTON, self.on_read_mac)

        self.mac_text_ctrl = wx.TextCtrl(self, value='MAC address', style=wx.TE_READONLY)
        empty_label = wx.StaticText(self, label='')

        flex_grid.AddMany([port_label, (serial_boxsizer, 1, wx.EXPAND),
                           file_label, (file_picker, 1, wx.EXPAND),
                           erase_label, (erase_boxsizer, 1, wx.EXPAND),
                           (read_mac_button, 1, wx.EXPAND), (self.mac_text_ctrl, 1, wx.EXPAND),
                           (empty_label, 1, wx.EXPAND), (upload_button, 1, wx.EXPAND)])
        flex_grid.AddGrowableRow(5, 1)
        flex_grid.AddGrowableCol(1, 1)

        hbox.Add(flex_grid, proportion=2, flag=wx.ALL | wx.EXPAND, border=15)
        self.SetSizer(hbox)

    def on_read_mac(self, event=None):
        print('Reading MAC')
        self.mac_text_ctrl.SetValue("")
        mac_add = esptool_read_mac(self._config.port)
        self.mac_text_ctrl.SetValue(mac_add)

    def on_upload(self, event):
        if self._config.port is None:
            print("no port selected")
            wx.MessageBox("No Port Selected !", caption="Select Port", style=wx.OK | wx.ICON_ERROR)
        elif self._config.firmware_path == "":
            print('no file is selected')
            wx.MessageBox("No file is selected !", caption="Select Firmware", style=wx.OK | wx.ICON_ERROR)
        else:
            print('Uploading...')
            print(self._config.port + ', ' + str(self._config.baud) + ", " + self._config.firmware_path)
            print('Erase flash = ' + str(self._config.erase_flash))
            self.on_read_mac()
            worker = EspToolThread(self, self._config, self.mac_text_ctrl)
            worker.start()

    def on_erase_change(self, event):
        rb = event.GetEventObject()

        if rb.GetLabel() == 'Yes':
            self._config.erase_flash = True
        else:
            self._config.erase_flash = False

        print('erase: ' + str(self._config.erase_flash))

    def on_pick_file(self, event):
        self._config.firmware_path = event.GetPath().replace("'", "")
        print('Firmware path: ' + self._config.firmware_path)

    def on_select_port(self, event):
        choice = event.GetEventObject()
        self._config.port = choice.GetString(choice.GetSelection())
        print("Port: " + self._config.port)

    def on_reload(self, event):
        print('port reload')
        self.choice.SetItems(self._get_serial_ports())

    @staticmethod
    def _get_serial_ports():
        ports = ["com1", "com2"]
        for port, desc, hwid in sorted(list_ports.comports()):
            ports.append(port)
        return ports


class EspFlasher(wx.Frame):
    def __init__(self, parent, title):
        super(EspFlasher, self).__init__(parent, title=title, size=(500, 400))
        # style=wx.DEFAULT_FRAME_STYLE | wx.NO_FULL_REPAINT_ON_RESIZE
        self.locale = wx.Locale(wx.LANGUAGE_ENGLISH)
        self.panel = MyPanel(self)


class MyApp(wx.App):
    def OnInit(self):
        self.SetAppName("ESP Flasher")
        frame = EspFlasher(parent=None, title='ESP Flasher')
        frame.Show()

        return True


def main():
    app = MyApp()
    app.MainLoop()


if __name__ == '__main__':
    __name__ = 'main'
    main()

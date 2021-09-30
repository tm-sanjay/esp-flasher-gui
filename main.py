import wx
import esptool
import threading
import serial
import sys
import os
from config_file import FlashConfig
from serial.tools import list_ports

DEVNULL = open(os.devnull, 'w')

__version__ = "0.0.3"
__auto_select__ = "Auto-select"
# TODO: auto-detect serial port(refer pyserial)


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
            argv.extend(["--baud", str(self._config.baud),
                         "--after", "hard_reset", "write_flash",
                         "--flash_size", "detect",
                         "--flash_mode", self._config.mode,
                         "0x00000", self._config.firmware_path])

            if self._config.erase_flash:
                argv.append('--erase-all')
            print(argv)
            print("Command: esptool.py %s\n" % " ".join(argv))

            # esptool.main(argv)

        except Exception as e:
            print("Unexpected error: {}".format(e))
            raise e

    # def read_mac(self):
    #     # read mac and update to UI
    #     self.mac = esptool_read_mac(self._config.port)
    #     wx.CallAfter(self.txt_ctrl.SetValue, self.mac)


#
# --------------------------------------------------
# GUI

class MyPanel(wx.Panel):
    filename = ''
    mac_address = ''

    def __init__(self, parent):
        super(MyPanel, self).__init__(parent)

        self._config = FlashConfig.load()

        # labels
        port_label = wx.StaticText(self, label='Serial Port')
        file_label = wx.StaticText(self, label='Firmware file')

        hbox = wx.BoxSizer(wx.HORIZONTAL)

        flex_grid = wx.FlexGridSizer(6, 2, 10, 10)

        self.choice = wx.Choice(self, choices=self._get_serial_ports())
        self.choice.Bind(wx.EVT_CHOICE, self.on_select_port)

        reload_button = wx.Button(self, label="Reload")
        reload_button.Bind(wx.EVT_BUTTON, self.on_reload)
        reload_button.SetToolTip("Reload serial device list")

        serial_boxsizer = wx.BoxSizer(wx.HORIZONTAL)
        serial_boxsizer.Add(self.choice, 1, wx.EXPAND)
        serial_boxsizer.Add(reload_button, flag=wx.LEFT, border=10)

        file_picker = wx.FilePickerCtrl(self, style=wx.FLP_USE_TEXTCTRL)
        file_picker.Bind(wx.EVT_FILEPICKER_CHANGED, self.on_pick_file)

        upload_button = wx.Button(self, label="Upload Firmware")
        upload_button.Bind(wx.EVT_BUTTON, self.on_upload)

        read_mac_button = wx.Button(self, label="Read MAC")
        read_mac_button.Bind(wx.EVT_BUTTON, self.on_read_mac)

        self.mac_text_ctrl = wx.TextCtrl(self, value='MAC address', style=wx.TE_READONLY)
        empty_label = wx.StaticText(self, label='')
        
        save_to_label = wx.StaticText(self, label='Save to Excel')

        auto_save_checkbox = wx.CheckBox(self, label="Auto Save")
        auto_save_checkbox.Bind(wx.EVT_CHECKBOX, self.on_auto_save)

        save_button = wx.Button(self, label="Save")
        save_button.Bind(wx.EVT_BUTTON, self.on_save)

        excel_boxsizer = wx.BoxSizer(wx.HORIZONTAL)
        excel_boxsizer.Add(auto_save_checkbox)
        excel_boxsizer.Add(save_button)

        flex_grid.AddMany([port_label, (serial_boxsizer, 1, wx.EXPAND),
                           file_label, (file_picker, 1, wx.EXPAND),
                           (read_mac_button, 1, wx.EXPAND), (self.mac_text_ctrl, 1, wx.EXPAND),
                           (empty_label, 1, wx.EXPAND), (upload_button, 1, wx.EXPAND),
                           (save_to_label, 1, wx.EXPAND), (excel_boxsizer, wx.EXPAND)])
        flex_grid.AddGrowableRow(5, 1)
        flex_grid.AddGrowableCol(1, 1)

        hbox.Add(flex_grid, proportion=2, flag=wx.ALL | wx.EXPAND, border=15)
        self.SetSizer(hbox)

    def on_read_mac(self, event=None):
        print('Reading MAC')
        if self._config.port is None:
            print("no port selected")
            wx.MessageBox("No Port Selected !", caption="Select Port", style=wx.OK | wx.ICON_ERROR)
        else:
            self.mac_text_ctrl.SetValue("")
            mac_add = esptool_read_mac(self._config.port)
            self.mac_text_ctrl.SetValue(mac_add)
            MyPanel.mac_address = mac_add
            # wx.MessageBox("Please reconnect the device or restart the App !", caption="Reconnect", style=wx.OK |
            # wx.ICON_WARNING)

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
            self.on_read_mac()
            worker = EspToolThread(self, self._config, self.mac_text_ctrl)
            worker.start()

    def on_pick_file(self, event):
        self._config.firmware_path = event.GetPath().replace("'", "")
        print('Firmware path: ' + self._config.firmware_path)
        MyPanel.filename = os.path.basename(self._config.firmware_path)
        print(MyPanel.filename)

    def on_select_port(self, event):
        choice = event.GetEventObject()
        self._config.port = choice.GetString(choice.GetSelection())
        print("Port: " + self._config.port)

    def on_reload(self, event):
        print('port reload')
        self.choice.SetItems(self._get_serial_ports())

    @staticmethod
    def _get_serial_ports():
        ports = []
        for port, desc, hwid in sorted(list_ports.comports()):
            ports.append(port)
        return ports

    def on_auto_save(self, event):
        print("on auto save")

    def on_save(self, event):
        print("on save")


class SettingsTab(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)

        self._config = FlashConfig.load()

        hbox = wx.BoxSizer(wx.HORIZONTAL)

        flex_grid = wx.FlexGridSizer(6, 1, 10, 10)

        # radio box for baud-rate selection
        baud_rate_list = ['9600', '57600', '74880', '115200', '230400', '460800', '921600']
        baud_box = wx.RadioBox()
        baud_box.Create(self, label='Baud Rate', choices=baud_rate_list,
                        majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        b_index = baud_rate_list.index(self._config.baud)
        # default value
        baud_box.SetSelection(b_index)
        baud_box.Bind(wx.EVT_RADIOBOX, self.on_baud_rate)

        # radio box for mode selection
        mode_list = ['qio', 'dio', 'dout']
        mode_box = wx.RadioBox()
        mode_box.Create(self, label='Mode', choices=mode_list,
                        majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        m_index = mode_list.index(self._config.mode)
        # default value
        mode_box.SetSelection(m_index)
        mode_box.Bind(wx.EVT_RADIOBOX, self.on_mode)

        save_button = wx.Button(self, label='Save')
        save_button.Bind(wx.EVT_BUTTON, self.on_save)

        # radio box for erasing flash
        erase_list = ['No', 'Yes']
        erase_box = wx.RadioBox()
        erase_box.Create(self, label='Erase Flash', choices=erase_list,
                         majorDimension=1, style=wx.RA_SPECIFY_ROWS)
        e_index = erase_list.index(self._config.erase_flash)
        # default value
        erase_box.SetSelection(e_index)
        erase_box.Bind(wx.EVT_RADIOBOX, self.on_erase)

        box1 = wx.BoxSizer(wx.HORIZONTAL)
        box1.Add(erase_box, flag=wx.RIGHT, border=10)
        box1.Add(mode_box, flag=wx.LEFT, border=10)

        flex_grid.AddMany([baud_box,
                           (box1, 1, wx.EXPAND),
                           save_button])
        flex_grid.AddGrowableRow(5, 1)
        flex_grid.AddGrowableCol(0, 1)

        hbox.Add(flex_grid, proportion=2, flag=wx.ALL | wx.EXPAND, border=15)
        self.SetSizer(hbox)

    def on_baud_rate(self, event):
        br = event.GetEventObject()
        print('On Baud Rate')
        self._config.baud = br.GetStringSelection()
        print(self._config.baud)

    def on_mode(self, event):
        m = event.GetEventObject()
        print('On Mode')
        self._config.mode = m.GetStringSelection()
        print(self._config.mode)

    def on_erase(self, event):
        e = event.GetEventObject()
        self._config.erase_flash = e.GetStringSelection()
        print('erase: ' + str(self._config.erase_flash))

    def on_save(self, event):
        print(f'mode:{self._config.mode}, baud rate:{self._config.baud},'
              f'erase:{self._config.erase_flash}')
        self._config.save()


class ExeclTab(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.output_file_path = ""

        hbox = wx.BoxSizer(wx.HORIZONTAL)

        flex_grid = wx.FlexGridSizer(6, 2, 10, 10)

        save_to_label = wx.StaticText(self, label="File Destination")
        dir_picker = wx.DirPickerCtrl(self, style=wx.FLP_USE_TEXTCTRL)
        dir_picker.Bind(wx.EVT_DIRPICKER_CHANGED, self.on_pick_dir)

        save_button = wx.Button(self, label="Save", pos=(20, 30))
        save_button.Bind(wx.EVT_BUTTON, self.on_save)

        flex_grid.AddMany([save_to_label, (dir_picker, 1, wx.EXPAND),
                           save_button])
        flex_grid.AddGrowableRow(5, 1)
        flex_grid.AddGrowableCol(1, 1)

        hbox.Add(flex_grid, proportion=2, flag=wx.ALL | wx.EXPAND, border=15)
        self.SetSizer(hbox)

    def on_save(self, event):
        from to_excel import Excel
        # Excel(path=self.output_file_path, mac_id=MyPanel.mac_address, file_name=MyPanel.filename)

    def on_pick_dir(self, event):
        self.output_file_path = event.GetPath()
        print(f'path: {self.output_file_path}')


class EspFlasher(wx.Frame):
    def __init__(self, parent, title):
        super(EspFlasher, self).__init__(parent, title=title, size=(500, 400))
        # style=wx.DEFAULT_FRAME_STYLE | wx.NO_FULL_REPAINT_ON_RESIZE
        self.locale = wx.Locale(wx.LANGUAGE_ENGLISH)

        notebook = wx.Notebook(self)

        tab1 = MyPanel(notebook)
        tab2 = SettingsTab(notebook)
        tab3 = ExeclTab(notebook)

        notebook.AddPage(tab1, "Main")
        notebook.AddPage(tab2, "Settings")
        notebook.AddPage(tab3, "Execl")

        # self.panel = MyPanel(self)
        # self._menu_bar()

    def _menu_bar(self):
        self.menuBar = wx.MenuBar()

        # File menu
        fileMenu = wx.Menu()
        item = fileMenu.Append(wx.ID_EXIT, '&Quit\tCtrl+Q')

        self.menuBar.Append(fileMenu, "&File")
        self.Bind(wx.EVT_MENU, self._on_exit, item)
        self.SetMenuBar(self.menuBar)

        # Settings menu
        settings = wx.Menu()
        item = settings.Append(wx.ID_NEW, 'Edit')

        self.menuBar.Append(settings, "&Settings")
        self.Bind(wx.EVT_MENU, self._on_settings, item)
        self.SetMenuBar(self.menuBar)

    def _on_exit(self, event):
        self.Close()

    def _on_settings(self, event):
        self.Close()


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

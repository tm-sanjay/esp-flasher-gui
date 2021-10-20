import pylightxl as xl
import os.path
from config_file import ExcelConfig
from datetime import date


class Excel:
    path = 'output.xlsx'
    saved_data = ''

    def __init__(self):
        self.db = xl.Database()
        self.output_file = 'output.xlsx'
        self.date = str(date.today())
        self._config = ExcelConfig.load()

    def set_output_path(self, path):
        Excel.path = path
        if not path:
            # file will be saved in same dir
            Excel.path = self.output_file
        else:
            Excel.path += '/' + self.output_file
        print(Excel.path)

    def save_data(self, mac_id, file_name):
        # check for file or create
        if os.path.isfile(Excel.path):
            self.db = xl.readxl(fn='output.xlsx')
            self.db.ws(ws="Sheet1")
        else:
            # add a blank worksheet to the db
            self.db.add_ws(ws="Sheet1")
            self.db.ws(ws='Sheet1').update_index(row=1, col=1, val='Sl-No')
            self.db.ws(ws='Sheet1').update_index(row=1, col=2, val='MAC-ID')
            self.db.ws(ws='Sheet1').update_index(row=1, col=3, val='Date')
            self.db.ws(ws='Sheet1').update_index(row=1, col=4, val='File Name')
            self._config.row_id = 2
            print('File does not exists, creating new Exel')

        # Column headings
        # sl number in col 1
        self.db.ws(ws='Sheet1').update_index(row=self._config.row_id, col=1, val=int(self._config.row_id - 1))
        # mac address in col 2
        self.db.ws(ws='Sheet1').update_index(row=self._config.row_id, col=2, val=mac_id)
        # date in col 3
        self.db.ws(ws='Sheet1').update_index(row=self._config.row_id, col=3, val=self.date)
        # filename in col 4
        self.db.ws(ws='Sheet1').update_index(row=self._config.row_id, col=4, val=file_name)
        self._config.row_id += 1

        # write out the db
        xl.writexl(db=self.db, fn=Excel.path)
        Excel.saved_data = f"{self._config.row_id - 2}, {mac_id}, {self.date}"
        print(f"\nExcel-saved: {Excel.saved_data}")
        self._config.save()

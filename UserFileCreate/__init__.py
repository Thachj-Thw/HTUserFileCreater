from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles.colors import Color, COLOR_INDEX



class Struct:
    def __init__(self) -> None:
        self.start_row = 5
        self.name_column = "B"
        self.room_column = "B"
        self.license_plate_car = "C"
        self.license_plate_moto = "D"
        self.license_plate_electric = "E"


class UserFileCreater:
    def __init__(self, file_path: str, struct: Struct = None) -> None:
        self.workbook = load_workbook(filename=file_path)
        self.sheet = self.workbook.active
        self.struct = struct if struct else Struct()
        self.save_file = Workbook()
        self.save_sheet = self.save_file.active
    
    def save(self, save_as: str = "output.xlsx") -> None:
        self.save_file.save(filename=save_as)
    
    def detect_room(self):
        for cell in self.sheet[self.struct.room_column]:
            fgColor = cell.fill.fgColor.rgb
            if type(fgColor) is str and fgColor != COLOR_INDEX[0]:
                # print(type(fgColor))
                print(cell.value)


class UserFile:
    def __init__(self) -> None:
        self.cells = []


class License_plate_Infomation:
    def __init__(self) -> None:
        self.user = ""
        self.room = ""
        self.number = ""
        self.type = ""

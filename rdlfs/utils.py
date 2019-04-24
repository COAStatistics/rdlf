import openpyxl
from openpyxl.utils import get_column_letter


class ExcelHandler:
    titles = {
        'sample': ['農戶編號', '調查姓名', '電話', '地址', '出生年', '原層別', '連結編號'],
        'household': ['出生年', '關係'],
        'crop_sbdy': ['[轉作補貼]', '項目', '作物名稱', '金額', '期別'],
        'disaster': ['[災害]', '項目', '災害', '核定作物', '核定面積', '金額'],
        '104_mon_hire_fir_half': ['[104年農普每月僱工]', '一月', '二月', '三月', '四月', '五月', '六月'],
        '104_mon_hire_sec_half': ['七月', '八月', '九月', '十月', '十一月', '十二月'],
        '106_shirt_hire_fir_half': ['一月', '二月', '三月', '四月', '五月', '六月'],
        '106_shirt_hire_sec_half': ['七月', '八月', '九月', '十月', '十一月', '十二月'],
    }

    def __init__(self):
        self.__col_index = 1
        self.__row_index = 1
        self.__wb = openpyxl.Workbook()
        self.__sheet = self.__wb.active

        self.__set_column_width()

    @property
    def column_index(self):
        return self.__col_index

    @column_index.setter
    def column_index(self, i):
        if i == -1:
            self.__col_index = 1
        else:
            self.__col_index += i

    @property
    def row_index(self):
        return self.__row_index

    @row_index.setter
    def row_index(self, i):
        if i == -1:
            self.__row_index = 1
        else:
            self.__row_index += i

    def __set_column_width(self):
        width = list(map(lambda x: x * 1.054, [14.29, 9.29, 16.29, 29.29, 9.29, 11.29, 11.29, 11.29, 11.29]))
        for i in range(1, len(width) + 1):
            self.__sheet.column_dimensions[get_column_letter(i)].width = width[i - 1]
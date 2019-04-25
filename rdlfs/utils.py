import openpyxl
from functools import reduce
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side


class ExcelHandler:
    titles = {
        'sample': ['農戶編號', '調查姓名', '電話', '地址', '出生年', '原層別', '連結編號'],
        'household': ['出生年', '關係'],
        'crop_sbdy': ['[轉作補貼]', '項目', '作物名稱', '期別'],
        'disaster': ['[災害]', '項目', '災害', '核定作物', '核定面積'],
        'crops_name': '[106y-107y作物]',
        'fir_half': ['[{}]', '一月', '二月', '三月', '四月', '五月', '六月'],
        'sec_half': ['七月', '八月', '九月', '十月', '十一月', '十二月'],
        '106y_hire': ['[106每月常僱員工]'],
    }

    def __init__(self):
        self.__col_index = 1
        self.__row_index = 1
        self.__wb = openpyxl.Workbook()
        self.__sheet = self.__wb.active
        self.__crop_set = set()

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
        self.row_index = 1

    def __set_title(self, title, start=1):
        for index, _title in enumerate(self.titles[title], start=start):
            self.__sheet.cell(column=index, row=self.row_index).value = _title
        self.row_index = 1

    def __set_sample_base_data(self, data):
        self.__set_title('sample')
        value_list = [data['farmer_num'], data['name'], data['tel'],
                      data['addr'], data['birthday'], data['layer'], data['link_num']]

        for index, val in enumerate(value_list, start=1):
            self.__sheet.cell(column=index, row=self.row_index).value = val
            self.__sheet.cell(column=index, row=self.row_index).alignment = Alignment(wrap_text=True)
        self.row_index = 1
        self.__sheet.cell(column=self.column_index, row=self.row_index).value = ' ' + '-'*206 + ' '
        self.row_index = 1

    def __set_household_data(self, members):
        self.__set_title('household')
        if not members:
            return
        for person in members:
            self.row_index = 1
            for index, val in enumerate(person, start=1):
                self.__sheet.cell(column=index, row=self.row_index).value = val
                self.__sheet.cell(column=index, row=self.row_index).alignment = Alignment(horizontal='left')

    def __set_crop_sbdy_data(self, crop_sbdy_list):
        if not crop_sbdy_list:
            return
        self.row_index = 2
        self.__set_title('crop_sbdy')
        self.__crop_set = crop_name = {i[0] for i in crop_sbdy_list}

        for index, val in enumerate(crop_name, start=1):
            if index >= 2:
                self.row_index = 1
            self.__sheet.cell(column=2, row=self.row_index).value = index
            self.__sheet.cell(column=3, row=self.row_index).value = val
            self.__sheet.cell(column=4, row=self.row_index).value = '1'

            if len(val) > 8:
                self.__sheet.cell(column=3, row=self.row_index).alignment = Alignment(wrap_text=True)

    def __set_disaster_data(self, disaster_list):
        if not disaster_list:
            return
        _dict = {}
        self.row_index = 2
        for d in disaster_list:
            key = (d[0], d[1])
            if key not in _dict:
                _dict[key] = float(d[2])
            else:
                _dict[key] = _dict.get(key) + float(d[2])
        self.__set_title('disaster')

        for index, _tuple in enumerate(_dict.items(), start=1):
            if index >= 2:
                self.row_index = 1
            self.__crop_set.add(_tuple[0][1])
            self.__sheet.cell(column=2, row=self.row_index).value = index
            self.__sheet.cell(column=3, row=self.row_index).value = _tuple[0][0]
            self.__sheet.cell(column=4, row=self.row_index).value = _tuple[0][1]
            self.__sheet.cell(column=5, row=self.row_index).value = _tuple[1]

    def __set_crops_name(self):
        self.row_index = 2
        crops = reduce(lambda a, b: a + ', ' + b, self.__crop_set)
        self.__sheet.cell(column=1, row=self.row_index).value = self.titles['crops_name']
        self.__sheet.cell(column=2, row=self.row_index).value = crops
        self.__crop_set.clear()

    def __set_104y__hire_or_106y_short_hire(self, hire_list, is104y=True):
        if not hire_list:
            return
        if is104y:
            self.titles['fir_half'][0].format('[104農普每月僱工]')
        else:
            self.titles['fir_half'][0].format('[106每月臨時僱工]')

        self.row_index = 2
        self.__set_title('fir_half')
        for index, mon, in enumerate(hire_list[:6]):
            self.__sheet.cell(column=2, row=self.row_index).value = mon

        self.row_index = 1
        self.__set_title('sec_half')
        for index, mon, in enumerate(hire_list[6:], start=2):
            self.__sheet.cell(column=2, row=self.row_index).value = mon

    def __set_106y_hire(self, hire_list):
        if not hire_list:
            return
        self.row_index = 2
        work_type = [d['工作類型'] for d in hire_list].insert(0, '工作類型')
        number = [d['常僱人數'] for d in hire_list].insert(0, '人數')
        mon = [d['months'] for d in hire_list].insert(0, '月份')
        self.__sheet.cell(column=1, row=self.row_index).value = self.titles['106y_hire']
        for _list in [work_type, number, mon]:
            for index, i in enumerate(_list, start=2):
                self.row_index = 1
                self.__sheet.cell(column=index, row=self.row_index).value = i

    def set_data(self, data):
        self.__set_sample_base_data(data)
        self.__set_household_data(data['household'])
        self.__set_crop_sbdy_data(data['crop_sbdy'])
        self.__set_crop_sbdy_data(data['disaster'])
        self.__set_crops_name()
        self.__set_104y__hire_or_106y_short_hire(data['mon_hire_104y'])
        self.__set_106y_hire(data['hire_106y'])
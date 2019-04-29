import datetime
import multiprocessing
import json
import os
import time
import utils
from generatedata import OUTPUT_DATA_DIR


def read_result_data(path) -> list:
    with open(path, encoding='utf8') as f:
        return json.loads(f.read())


def write_data_to_excel(path_tuple) -> None:
    excel_path, json_path = path_tuple
    county = json_path[json_path.index('.')-3:json_path.index('.')]
    data_list = read_result_data(json_path)
    handler = utils.ExcelHandler(county, excel_path)
    for data in data_list:
        handler.set_data(data)
    handler.save()

if __name__ == '__main__':
    start_time = time.time()
    json_path = os.path.join(OUTPUT_DATA_DIR, 'json/separate_json')
    excel_path = os.path.join(OUTPUT_DATA_DIR, '公務資料/' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '_公務資料(縣市切割)')
    if not os.path.isdir(excel_path):
        os.mkdir(excel_path)
    
    _list = []
    for file_name in os.listdir(json_path):
        _list.append((excel_path, os.path.join(json_path, file_name)))
    
    with multiprocessing.Pool(4) as p:
        [p.map(write_data_to_excel, _list)]
        
    m, s = divmod(time.time() - start_time, 60)
    print(int(m), 'min', round(s, 1), 'sec')
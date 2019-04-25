import datetime
import json
import os
import time
import utils
from generatedata import FILES, OUTPUT_DATA_DIR


def read_result_data() -> dict:
    data_dict = {}
    for data in json.loads(open(FILES['result_json'], encoding='utf8').read()).values():
        county = data.get('addr')[:3]
        if county not in data_dict:
            data_dict[county] = [data]
        else:
            data_dict.get(county).append(data)

    return data_dict


def write_data_to_excel(path) -> None:
    data_dict = read_result_data()
    count = 0
    for county, data_set in data_dict.items():
        handler = utils.ExcelHandler(county, path)
        for data in data_set:
            handler.set_data(data)
        count += 1
        print(count, '/', len(data_dict), '(' + county + ')', '...')


if __name__ == '__main__':
    start_time = time.time()
    path = os.path.join(OUTPUT_DATA_DIR, '公務資料/' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '_公務資料(縣市切割)')
    if not os.path.isdir(path):
        os.mkdir(path)
    write_data_to_excel(path)
    m, s = divmod(time.time() - start_time, 60)
    print(int(m), 'min', round(s, 1), 'sec')
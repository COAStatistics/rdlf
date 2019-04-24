import json
import os
import re
import time
from collections import namedtuple
from collections import OrderedDict
from dbconn import DatabaseConnection
from log import log, err_log

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

INPUT_DATA_DIR = os.path.join(BASE_DIR, 'input/107勞動力調查')

OUTPUT_DATA_DIR = os.path.join(BASE_DIR, 'output/107勞動力調查')

FILES = {
    'samples': os.path.join(INPUT_DATA_DIR, '104農普勞動力名冊.json'),
    'lack': os.path.join(INPUT_DATA_DIR, '106勞動力常缺.json'),
    'hire': os.path.join(INPUT_DATA_DIR, '106勞動力常僱.json'),
    'short_lack': os.path.join(INPUT_DATA_DIR, '106勞動力臨缺.json'),
    'short_hire': os.path.join(INPUT_DATA_DIR, '106勞動力臨僱.json'),
    'households': os.path.join(INPUT_DATA_DIR, 'coa_stat_d03_10804.txt'),
    'result_json': os.path.join(OUTPUT_DATA_DIR, 'json/data.json'),
}

# defined namedtuple attribute
PERSON_ATTR = ['addr_code', 'id', 'birthday', 'household_num', 'addr', 'role', 'annotation', 'h_type', 'h_code', ]

# use namedtuple promote the readable and flexibility of code
Person = namedtuple('Person', PERSON_ATTR)

monthly_employee_dict = {}
insurance_data = {}

all_samples = json.loads(open(FILES['samples'], encoding='utf8').read())
hire_106y_dict = {}
short_hire_106y_dict = {}
lack_106y_dict = {}
short_lack_106y_dict = {}
households = {}
result_data = {}


def data_calssify() -> None:
    # 有效身分證之樣本
    valid_samples_id_dict = get_valid_samples_id()

    # 樣本與戶號對照 dict
    # key: 樣本身分證字號, value: 樣本戶號
    comparison_dict = {}

    with open(FILES['households'], 'r', encoding='utf8') as f:
        for coa_data in f:

            # create Person object
            person = Person._make(coa_data.strip().split(','))
            pid = person.id
            hhn = person.household_num

            # 以戶號判斷是否存在, 存在則往列表(戶內人口列表)裡增加成員, 否則新增一戶
            if hhn in households:
                households.get(hhn).append(person)
            else:
                # 戶內人口列表
                persons = [person]
                households[hhn] = persons

            # 樣本 ID 若能對應到有效 ID dict, 就往對照 dict 裡新增(key: id, value: 戶號)
            if pid in valid_samples_id_dict:
                comparison_dict[pid] = hhn
    init_data(comparison_dict)


def get_valid_samples_id() -> dict:
    """
    讀取樣本 JSON 檔並迭代撿查 ID 是否重複且有效
    並將錯誤記錄至 log 檔

    :return valid_id_dict: 不重複且有效的樣本ID字典，value 為空值
    """
    no_id_count = 0
    duplicate_count = 0
    valid_id_dict = {}

    for sample in all_samples:
        # 去除重複的人與檢查身份證字號格式是否正確
        if sample['id'] not in valid_id_dict and re.match('^[A-Z]{1}[1-2]{1}[0-9]{8}$', sample['id']):
            valid_id_dict[sample['id']] = ''
        else:
            err_log.error('sample name = ', sample['name'], ', sample id = ', sample['id'])

            if sample['id'] == '0':
                no_id_count += 1
            else:
                duplicate_count += 1

    log.info('no id count = ', no_id_count, ', duplicate count = ', duplicate_count)
    global sample_count
    sample_count = len(all_samples)
    return valid_id_dict


def get_members_base_data(members) -> list:
    """
    迭代整戶的人(每個人都是一個 Person 物件) 只取得出生年與戶長關係並重新回傳一個新列表
    :param members: 列表(整戶的人的資料)
    :return members_data_list: 巢状列表，裡面每個列表都是一個成員的資料
    """

    # 一整戶的列表
    members_data_list = []

    # 每個成員都是 Person 物件
    for person in members:
        person_data = [str(int(person.birthday[:3])), person.role]
        members_data_list.append(person_data)

    return members_data_list


def get_data_set(members) -> dict:
    """
    迭代整戶的人(每個人都是一個 Person 物件) 根據年齡來訪問資料庫，以減少時間成本
    以 ID 為參數去資料庫查詢休耕轉作補貼資料與天然災害資料
    :param members: 列表(整戶的人的資料)
    :return dict: 包含轉作與天災的字典
    """

    # 大列表，裡面每個列表都是一筆資料
    disaster_list = []
    crop_sbdy_list = []
    db = DatabaseConnection.get_db_instance()

    for person in members:
        age = 107 - int(person.birthday[:3])
        if age >= 18:
            DatabaseConnection.set_pid(person.id)
            disaster = db.get_disaster()
            if disaster:
                disaster_list.extend(disaster)
                log.info(person.id, ', disaster = ', disaster_list)

            crop_sbdy = db.get_crop_subsidy()
            if crop_sbdy:
                crop_sbdy_list.extend(crop_sbdy)
                log.info(person.id, ', crop_sbdy = ', crop_sbdy_list)

    return {'disaster': disaster_list, 'crop_sbdy': crop_sbdy_list}


def get_104_month_hire(sample) -> list:
    mon_hire_list = [
        sample["hire_Jan"],
        sample["hire_Feb"],
        sample["hire_March"],
        sample["hire_April"],
        sample["hire_May"],
        sample["hire_June"],
        sample["hire_July"],
        sample["hire_Aug"],
        sample["hire_Sep"],
        sample["hire_Oct"],
        sample["hire_Nov"],
        sample["hire_Dec"],
    ]
    return mon_hire_list


def get_106_hire_or_lack(farmer_num, name):
    """
    因 常僱 常缺 臨缺 JSON 檔裡的資料格式一樣，因此併成一個 fn 來寫，需再傳入檔名
    :param farmer_num: 農戶編號
    :param name: 檔名，用來判斷選擇變數
    :return list or empty list: 若有對應到農戶編號則回傳列表，否則回傳空列表
    """
    # 判斷要開啟哪個檔
    choose = {
        'hire': (hire_106y_dict, FILES['hire']),
        'lack': (lack_106y_dict, FILES['lack']),
        'short_lack': (short_lack_106y_dict, FILES['short_lack']),
    }
    _dict, _name = choose[name]

    # lazy read file, 如果沒資料才讀取
    # key: farmer_num, value: list(月份資料)，同個農戶可能有多筆資料
    if not _dict:
        for d in json.loads(open(_name, encoding='utf8').read()):
            if d['農戶編號'] not in _dict:
                _dict[d['農戶編號']] = [d]
            else:
                _dict.get(d['農戶編號']).append(d)

    if farmer_num in _dict:
        return _dict.get(farmer_num)
    else:
        return []


def get_106_short_hire(farmer_num):
    if not short_hire_106y_dict:
        for d in json.loads(open(FILES['short_hire'], encoding='utf8').read()):
            if d['農戶編號'] not in short_hire_106y_dict:
                short_hire_106y_dict[d['農戶編號']] = [
                    d["Jan"],
                    d["Feb"],
                    d["Mar"],
                    d["Apr"],
                    d["May"],
                    d["Jun"],
                    d["Jul"],
                    d["Aug"],
                    d["Sep"],
                    d["Oct"],
                    d["Nov"],
                    d["Dec"],
                ]
    if farmer_num in short_hire_106y_dict:
        return short_hire_106y_dict.get(farmer_num)
    else:
        return []


def init_data(comparison_dict) -> None:
    total = len(all_samples)
    count = 0

    for sample in all_samples:
        birthday = ''
        members_data = []
        data_set = {}

        if sample['id'] in comparison_dict:
            members = households.get(comparison_dict[sample['id']])
            birthday = [int(i.birthday[:3]) for i in members if i.id == sample['id']].pop()
            members_data = get_members_base_data(members)
            data_set = get_data_set(members)

        mon_hire_104y_list = get_104_month_hire(sample)
        hire_106y_list = get_106_hire_or_lack(sample['farmer_num'], 'hire')
        short_hire_106y_list = get_106_short_hire(sample['farmer_num'])
        lack_situation = sample.get('lacks106')
        lack_106y_list = get_106_hire_or_lack(sample['farmer_num'], 'lack')
        short_lack_106y_list = get_106_hire_or_lack(sample['farmer_num'], 'short_lack')

        generate_json_data(sample, birthday, members_data, data_set, mon_hire_104y_list,
                           hire_106y_list, short_hire_106y_list, lack_situation, lack_106y_list, short_lack_106y_list)

        count += 1
        print('{} / {} ...'.format(count, total))

    db = DatabaseConnection.get_db_instance()
    db.close_conn()
    output_json_data()


def generate_json_data(sample, birthday, household, data_set, mon_hire_104y,
                       hire_106y, short_hire_106y, lack_situation, lack_106y, short_lack_106y):
    data = OrderedDict()
    data['farmer_num'] = str(sample['farmer_num'])
    data['id'] = sample['id']
    data['name'] = sample['name']
    data['tel'] = sample['tel']
    data['addr'] = sample['addr']
    data['birthday'] = str(birthday)
    data['layer'] = str(sample['strata'])
    data['link_num'] = str(sample['link_num'])
    data['household'] = household
    data['crop_sbdy'] = data_set.get('crop_sbdy', [])
    data['disaster'] = data_set.get('disaster', [])
    data['mon_hire_104y'] = mon_hire_104y
    data['hire_106y'] = hire_106y
    data['short_hire_106y'] = short_hire_106y
    data['lack_situation'] = lack_situation
    data['lack_106y'] = lack_106y
    data['short_lack_106y'] = short_lack_106y

    result_data[sample['farmer_num']] = data


def output_json_data() -> None:
    with open(FILES['result_json'], 'w', encoding='utf8') as f:
        f.write(json.dumps(result_data, ensure_ascii=False))
    print('complete', len(result_data), ' records')
    log.info(len(result_data), ' records')


if __name__ == '__main__':
    start_time = time.time()
    data_calssify()
    m, s = divmod(time.time() - start_time, 60)
    print(int(m), 'min', round(s, 1), 'sec')
    log.info(int(m), ' min ', round(s, 1), ' sec')

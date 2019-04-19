import json
import os
import xlrd
import re
import time
from collections import namedtuple
from collections import OrderedDict
from dbconn import DatabaseConnection
from log import log, err_log


BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

INPUT_DATA_DIR = os.path.join(BASE_DIR, 'input/107勞動力調查')

FILES = {
    'samples': os.path.join(INPUT_DATA_DIR, '104農普勞動力名冊.json'),
    'lack': os.path.join(INPUT_DATA_DIR, '106勞動力常缺.json'),
    'hire': os.path.join(INPUT_DATA_DIR, '106勞動力常僱.json'),
    'lack_short': os.path.join(INPUT_DATA_DIR, '106勞動力臨缺.json'),
    'hire_short': os.path.join(INPUT_DATA_DIR, '106勞動力臨僱.json'),
    'households': os.path.join(INPUT_DATA_DIR, 'coa_stat_d03_10804.txt'),
    }

OUTPUT_PATH = '..\\..\\output\\json\\公務資料.json'
THIS_YEAR = 107
ANNOTATION_DICT = {'0': '', '1': '死亡', '2': '除戶'}
DEAD_LIST = []

# defined namedtuple attribute
PERSON_ATTR = ['addr_code','id', 'birthday', 'household_num', 'addr', 'role', 'annotation', 'h_type', 'h_code', ]

# use namedtuple promote the readable and flexibility of code
Person = namedtuple('Person', PERSON_ATTR)

monthly_employee_dict = {}
insurance_data = {}

all_samples = json.loads(open(FILES['samples'], encoding='utf8').read())
households = {}
official_data = {}
sample_count = 0


def load_monthly_employee() -> None:
    sample_list = [line.strip().split('\t') for line in open(MON_EMP_PATH, 'r', encoding='utf8')]
    global monthly_employee_dict; monthly_employee_dict = {sample[0].strip() : sample[1:] for sample in sample_list} #Key is farmer id


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
                persons = []
                # 往列表裡添加成員
                persons.append(person)
                households[hhn] = persons
                
            # 樣本 ID 若能對應到有效 ID dict, 就往對照 dict 裡新增(key: id, value: 戶號)
            if pid in valid_samples_id_dict:
                comparison_dict[pid] = hhn
    build_official_data(comparison_dict)


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
    global sample_count; sample_count = len(all_samples)
    return valid_id_dict


def build_official_data(comparison_dict) -> None:
    no_hh_count = 0
    count = 0
    db = DatabaseConnection()
    person_key = ['birthday', 'role', 'annotation', 'farmer_insurance', 'elder_allowance', 'national_pension',
                  'labor_insurance', 'labor_pension', 'farmer_insurance_payment', 'scholarship', 'sb']
    #key dict: for readable
    k_d = {person_key[i]:i for i in range(len(person_key))}

    for sample in all_samples:
        count += 1
        address, birthday, farmer_id, farmer_num = '', '', '', ''
        # json 資料
        json_data = OrderedDict()
        json_household = []
        json_disaster = []
        json_crop_sbdy = []
        
        farmer_id = sample.id
        farmer_num = sample.num
        if farmer_id in comparison_dict:
            household_num = comparison_dict.get(farmer_id)
            if household_num in households:
                
                # households.get(household_num) : 每戶 
                # person : 每戶的每個人
                # person is a Person object
                for person in households.get(household_num):
                    name = ''
                    pid = person.id
                    
                    # json data 主要以 sample 的人當資料，所以要判斷戶內人是否為 sample
                    if pid == sample.id:
                        name = sample.name
                        address = sample.addr
                        # 民國年
                        birthday = str(int(person.birthday[:3]))
                        
                    # 轉成實際年齡
                    age = THIS_YEAR - int(person.birthday[:3])
                    DatabaseConnection.pid = pid
                        
                    # json 裡的 household 對應一戶裡的所有個人資料
                    json_hh_person = [''] * 11
                    
                    json_hh_person[k_d['birthday']] = str(int(person.birthday[:3]))
                    json_hh_person[k_d['role']] = person.role
                    annotation = ANNOTATION_DICT.get(person.annotation)
                    if annotation == '死亡':
                        DEAD_LIST.append(person.id)
                    json_hh_person[k_d['annotation']] = annotation
                    
                
                    # 根據年齡來過濾是否訪問 db
                    # 農保至少15歲
                    if age >= 15:
                        json_hh_person[k_d['farmer_insurance']] = db.get_farmer_insurance()
                        # 老農津貼至少65歲
                        if age >= 65:
                            json_hh_person[k_d['elder_allowance']] = db.get_elder_allowance()
                        # 佃農18-55歲，地主至少18歲
                        if age >= 18:
                            json_hh_person[k_d['sb']] = db.get_landlord()
                            if age <= 55:
                                json_hh_person[k_d['sb']] += db.get_tenant_farmer()
                            subsidy = [
                                    name,
                                    db.get_tenant_transfer_subsidy(),
                                    db.get_landlord_rent(),
                                    db.get_landlord_retire()
                                ]
                            if any((i != '0') for i in subsidy[1:]):
                                json_sb_sbdy.append(subsidy)
                                log.info(pid, ', sbSbdy = ', json_sb_sbdy)
                                
                            disaster = db.get_disaster()
                            if disaster:
                                json_disaster.extend(disaster)
                                log.info(pid, ', disaster = ', json_disaster)
                                
                            declaration = db.get_declaration()
                            if declaration and declaration not in json_declaration:
                                json_declaration += declaration + ','
                                assert len(json_declaration) != 0
                                log.info(pid, ', declaration = ', json_declaration)
                                
                            crop_sbdy = db.get_crop_subsidy()
                            if crop_sbdy:
                                json_crop_sbdy.extend(crop_sbdy)
                                log.info(pid, ', crop_sbdy = ', json_crop_sbdy)
                    json_household.append(json_hh_person)
        else:
            DatabaseConnection.pid = farmer_id
            address = sample.addr
            json_hh_person = [''] * 11
            if sample.id in insurance_data:
                insurance = insurance_data.get(sample.id)
                for i in range(5, 9):
                    json_hh_person[i] = format(insurance[i-5], '8,d')
                    
            json_household.append(json_hh_person)
            if sample.id:
                no_hh_count += 1
                err_log.error(no_hh_count, ', Not in household file. ', sample)
            
        # create json data
        json_data['name'] = sample.name
        json_data['address'] = address
        json_data['birthday'] = birthday
        json_data['farmerId'] = farmer_id
        json_data['telephone'] = sample.tel
        json_data['layer'] = sample.layer
        json_data['serial'] = farmer_num[-5:]
        json_data['household'] = json_household
        json_data['monEmp'] = monthly_employee_dict.get(farmer_num, [])
        json_data['cropSbdy'] = json_crop_sbdy
        json_data['disaster'] = json_disaster
        
        official_data[farmer_num] = json_data
        print('%.2f%%' %(count/sample_count * 100))
    db.close_conn()
    output_josn(official_data)
    

def output_josn(data) -> None:
    with open(OUTPUT_PATH, 'w', encoding='utf8') as f:
        f.write(json.dumps(data,  ensure_ascii=False))
    print('complete', len(official_data), ' records')
    log.info(len(official_data), ' records')


if __name__ == '__main__' :
    print(BASE_DIR, INPUT_DATA_DIR)
    data = json.loads(open(FILES['samples'], encoding='utf8').read())
    data_calssify()
    
# start_time = time.time()
# load_insurance()
# data_calssify()
# m, s = divmod(time.time()-start_time, 60)
# print(int(m), 'min', round(s, 1), 'sec')
# log.info(int(m), ' min ', round(s, 1), ' sec')
#  
# with open('..\\..\\output\\dead.txt', 'w') as f:
#     for i in DEAD_LIST:
#         f.write(i + '\n')

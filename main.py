import random
import copy
import string
import pandas as pd
import data_generator as dg
from excel_utils import find_table_start, read_table_data, write_data_to_excel
from json_utils import write_data_to_json

not_same_data_buffer = {}
prefix_data_buffer = {}
suffix_data_buffer = {}
xlsx_file = 'data.xlsx'
data_test_col = 0

def add_not_same_properties(ppt_data_dict):
    global not_same_data_buffer
    data_buffer_local = not_same_data_buffer
    if ppt_data_dict["Variable"] not in data_buffer_local:
        data_buffer_local[ppt_data_dict["Variable"]] = []
    not_same_data_buffer = data_buffer_local

def add_prefix_properties(ppt_data_dict, prefix):
    global prefix_data_buffer
    data_buffer_local = prefix_data_buffer
    if ppt_data_dict["Variable"] not in data_buffer_local:
        data_buffer_local[ppt_data_dict["Variable"]] = prefix
    prefix_data_buffer = data_buffer_local

def add_suffix_properties(ppt_data_dict, suffix):
    global suffix_data_buffer
    data_buffer_local = suffix_data_buffer
    if ppt_data_dict["Variable"] not in data_buffer_local:
        data_buffer_local[ppt_data_dict["Variable"]] = suffix
    suffix_data_buffer = data_buffer_local

def not_same_properties_handler(data):
    global not_same_data_buffer
    new_value = ""
    for key in not_same_data_buffer.keys():
        if key not in data:
            continue
        if data[key] not in not_same_data_buffer[key]:
            continue
        if not is_number(data[key]):
            new_value = get_unique_string(not_same_data_buffer[key])
        else:
            new_value = dg.get_unique_random_string(not_same_data_buffer[key], dg.random_digit_string, len(data[key]))
        not_same_data_buffer[key].append(new_value)
        data[key] = new_value

def suffix_properties_handler(data):
    global suffix_data_buffer
    new_value = ""

    for element in suffix_data_buffer.keys():
        if element not in data:
            continue
        new_value = data[element] + suffix_data_buffer[element]
        data[element] = new_value

def prefix_properties_handler(data):
    global prefix_data_buffer
    new_value = ""

    for element in prefix_data_buffer.keys():
        if element not in data:
            continue
        new_value = prefix_data_buffer[element] + data[element]
        data[element] = new_value

def properties_handler(prop_strip, ppt_data_dict):
    key1 = "Tiền tố: "
    key2 = "Hậu tố: "
    if prop_strip == "Không trùng":
        add_not_same_properties(ppt_data_dict)
    elif key1 in prop_strip:
        # Giả sử prop_strip có dạng "Tiền Tố: @gmail.com"
        prefix = prop_strip.split(key1)[-1]

        add_prefix_properties(ppt_data_dict, prefix)
    elif key2 in prop_strip:
        # Giả sử prop_strip có dạng "Hậu Tố: @example.com"
        suffix = prop_strip.split(key2)[-1]
        add_suffix_properties(ppt_data_dict, suffix)
    elif prop_strip == "":
        #nothing
        pass
    else:
        print(f"Properties {prop_strip} không hợp lệ.")

def init_properties_data():
    file_path = xlsx_file
    ppt_data = read_table_data(file_path, header_fill_color="FFFFC000", sheetname="Properties")
    for index, row in ppt_data.iterrows():
        ppt_data_dict = row.to_dict()
        properties_list = ppt_data_dict['Properties'].split(',')  # Tách theo dấu phẩy
        for prop in properties_list:
            prop_strip = prop.strip()
            properties_handler(prop_strip, ppt_data_dict)

def init_data():
    file_path = xlsx_file  # Thay thế đường dẫn đến file Excel của bạn
    df = read_table_data(file_path, header_fill_color="FFFFC000", sheetname="Initialize")
    # Chuyển đổi bảng dữ liệu thành một dictionary với cột "Variable" là key
    data_dict = pd.Series(df['Default Value'].values, index=df['Variable']).to_dict()
    return data_dict

def read_testcase_data():
    global data_test_col
    file_path = xlsx_file  # Thay thế đường dẫn đến file Excel của bạn
    df = read_table_data(file_path, header_fill_color="FFFFC000", sheetname="Testcase")
    if 'Data Test' in df.columns:
        # Lấy chỉ số cột của "Data Test"
        data_test_col = df.columns.get_loc('Data Test')
        print(data_test_col)
    else:
        print("Không tìm thấy cột 'Data Test' trong bảng dữ liệu.")

    return df

def action_handler(data, modified_ele, action, length=0):
    if modified_ele not in data:
        return
    res = ""
    string_buffer = []
    if modified_ele in not_same_data_buffer:
        string_buffer = not_same_data_buffer[modified_ele]
    match action:
        case "Random chữ cái":
            res = dg.get_unique_random_string(string_buffer, dg.random_letter_string, length)
        case "Random số":
            res = dg.get_unique_random_string(string_buffer, dg.random_digit_string, length)
        case "Random kí tự đặc biệt":
            res = dg.get_unique_random_string(string_buffer, dg.random_special_characters, length)
        case "Nhập trùng":
            data[modified_ele] = random.choice(not_same_data_buffer[modified_ele]) if modified_ele in not_same_data_buffer and not_same_data_buffer[modified_ele] else data[modified_ele]
            return
        case "Không nhập":
            data[modified_ele] = ""
            return
        case "Không gửi":
            del data[modified_ele]
            return
        case _: 
             print("Hành động không hợp lệ.")
             return
    if modified_ele in prefix_data_buffer:
        res += prefix_data_buffer[modified_ele]
    if modified_ele in suffix_data_buffer:
        res = prefix_data_buffer[modified_ele] + res
    data[modified_ele] = res

# Kiểm tra xem một giá trị có phải là số hay không
def is_number(value):
    try:
        # Kiểm tra xem giá trị có thể chuyển thành số hay không
        float(value)
        return True
    except ValueError:
        return False

# Hàm tạo giá trị duy nhất
def get_unique_string(existing_values):
    new_value = None
    counter = 1
    
    while True:
        # Tạo giá trị mới dựa trên kiểu giá trị hiện tại (ở đây là chuỗi số + số đếm)
        new_value = ''.join(random.choices(string.ascii_letters + string.digits, k=8)) + str(counter)
        
        # Nếu giá trị này chưa tồn tại trong existing_values, trả về giá trị này
        if new_value not in existing_values:
            break
        
        counter += 1
    
    return new_value

def save_data(data):
    global data_buffer
    for key in not_same_data_buffer.keys():
        if key not in data:
            continue
        not_same_data_buffer[key].append(data[key])

def modify_data_base_on_properties(data):
    not_same_properties_handler(data)
    prefix_properties_handler(data)
    suffix_properties_handler(data)

def write_data_test_to_excel(row, col, data, sheet_name):
    file_path = xlsx_file
    write_data_to_excel(xlsx_file, row, col, data, sheet_name)

def main():
    # Đọc đường dẫn tới file Excel
    global data_test_col
    json_data = []
    init_properties_data()
    default_data = init_data()
    testcase_data = read_testcase_data()
    data_test_col_local = data_test_col
    for index, row in testcase_data.iterrows():
        out_data = copy.deepcopy(default_data)
        modify_data_base_on_properties(out_data)
        tc_dict = row.to_dict()
        modified_ele = tc_dict["Variable"]
        action_handler(out_data, modified_ele, tc_dict["Action"], tc_dict["Length"])
        write_data_test_to_excel(index, data_test_col_local, out_data, "Testcase")
        out_data["Method"] = tc_dict["Method"]
        out_data["Expect Result"] = tc_dict["Expect Result"]
        save_data(out_data)
        json_data.append(out_data)
    write_data_to_json(json_data, 'data.json')

# Chạy hàm main khi file này được thực thi trực tiếp
if __name__ == "__main__":
    main()
import json

def write_data_to_json(data, output_file_path):
    """
    Ghi dữ liệu vào một tệp JSON.

    Parameters:
    - data: Dữ liệu cần ghi (dạng dictionary hoặc list).
    - output_file_path: Đường dẫn đến tệp JSON cần ghi dữ liệu.

    """
    try:
        with open(output_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data, json_file, ensure_ascii=False, indent=4)
        print(f"Dữ liệu đã được ghi vào {output_file_path} thành công.")
    except Exception as e:
        print(f"Lỗi khi ghi dữ liệu vào tệp: {e}")
# data_generator.py

import random
import string

def random_digit_string(length=0):
    """
    Tạo chuỗi số ngẫu nhiên có độ dài "length".
    """
    return ''.join(random.choice('0123456789') for _ in range(length))

def random_letter_string(length=0): 
    """
    Tạo chuỗi ký tự ngẫu nhiên có độ dài "length" gồm chữ cái.
    """
    return ''.join(random.choice(string.ascii_letters) for _ in range(length))

def random_special_characters(length=10):
    """
    Tạo chuỗi ký tự đặc biệt ngẫu nhiên có độ dài "length".
    """
    return ''.join(random.choice(string.punctuation) for _ in range(length))

def get_unique_random_string(generated_strings, generator_func=None, length=10):
    """
    Tạo chuỗi ngẫu nhiên chưa có trong tập hợp "generated_strings".
    """
    result = "None"
    # Sử dụng hàm tạo chuỗi ngẫu nhiên được truyền vào hoặc mặc định
    if generator_func is None:
        generator_func = random_letter_string
    
    while True:
        random_string = generator_func(length)
        if random_string not in generated_strings:
            result = random_string
            break
    return result

def get_random_available_string(generated_strings):
    """
    Chọn ngẫu nhiên một chuỗi đã được tạo từ tập hợp "generated_strings".
    """
    if not generated_strings:
        return ""
    return random.choice(list(generated_strings))
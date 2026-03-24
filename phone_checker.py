import re

def is_phone_number(text):
    """
    判断输入内容是否为电话号码
    支持中国大陆手机号和固定电话格式
    """
    # 去除空格和横线
    text = text.strip().replace(' ', '').replace('-', '')

    # 手机号：1开头，共11位
    mobile_pattern = r'^1[3-9]\d{9}$'

    # 固定电话：区号(3-4位) + 号码(7-8位)
    # 格式：01012345678 或 075512345678
    landline_pattern = r'^0\d{2,3}\d{7,8}$'

    if re.match(mobile_pattern, text) or re.match(landline_pattern, text):
        return True
    return False


if __name__ == "__main__":
    while True:
        user_input = input("请输入内容：")
        if is_phone_number(user_input):
            print(f"'{user_input}' 是有效的电话号码")
            break
        else:
            print(f"'{user_input}' 不是有效的电话号码，请重新输入！")

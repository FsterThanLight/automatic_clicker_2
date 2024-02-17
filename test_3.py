import random
import string


def generate_random_alphanumeric(length):
    # 生成随机字母和数字的组合
    characters = string.ascii_letters + string.digits
    # 从字符集中随机选择字符，重复 length 次，并将结果连接成字符串
    return ''.join(random.choice(characters) for _ in range(length))


# 生成长度为 10 的随机字母和数字串
random_string = generate_random_alphanumeric(10)
print(random_string)

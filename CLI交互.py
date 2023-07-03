import os
import re
import csv
import docx
import datetime
from copy import deepcopy
import logging

# 配置日志
log_folder_name = "logging"
log_folder_path = os.path.join(os.getcwd(), log_folder_name)
os.makedirs(log_folder_path, exist_ok=True)
log_file_name = log_folder_name + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".log"
log_file_path = os.path.join(log_folder_path, log_file_name)



log_file_path = os.path.join(log_folder_path, log_file_name)

logging.basicConfig(filename=log_file_path, level=logging.DEBUG, format='[%(asctime)s] [%(levelname)s] [%(module)s/%(funcName)s] %(message)s')

# 正则表达式规则
folder_name_regex = re.compile(r'^[\u4E00-\u9FFFA-Za-z\d\-_.]{1,255}$')
filename_regex = re.compile(r'[^A-Za-z\d_.\u4E00-\u9FFF-]|(?!(.docx?|doc)$)')

# 创建保存Word文档的文件夹
default_folder_name = "Excel-Cvs-Python-Docx"
use_default_folder = input("是否使用默认的文件夹名称（用于储存处理后的Word文档）？（y/n）：") == "y"

if use_default_folder:
    folder_name = default_folder_name
else:
    folder_name = input("请输入文件夹名称：")
    while not folder_name_regex.match(folder_name):
        logging.warning("[WARNING] [create_word_folder] 文件夹名称不合法，请重新输入。")
        folder_name = input("请输入文件夹名称：")

# 获取文件夹存储路径
default_folder_path = os.getcwd()
use_default_path = input("是否使用默认的文件夹存储路径？（即代码所在路径）（y/n）：") == "y"

if not use_default_path:
    folder_path = input("请输入文件夹存储路径：")
    while not os.path.exists(folder_path):
        logging.warning("[WARNING] [get_folder_path] 路径不存在，请重新输入。")
        folder_path = input("请输入文件夹存储路径：")
else:
    folder_path = default_folder_path

# 创建文件夹
folder_path = os.path.join(folder_path, folder_name)
os.makedirs(folder_path, exist_ok=True)
print(f"已创建文件夹：{folder_name}")
logging.info("[INFO] [create_folder] 已创建文件夹：" + folder_name)

# 获取CSV文件路径
default_csv_path = default_folder_path
use_default_csv_path = input("是否使用默认的CSV文件存储位置？（即代码所在路径）（y/n）：") == "y"

if not use_default_csv_path:
    csv_path = input("请输入CSV文件存储路径：")
    while not os.path.exists(csv_path):
        logging.warning("[WARNING] [get_csv_path] 路径不存在，请重新输入。")
        csv_path = input("请输入CSV文件存储路径：")
else:
    csv_path = default_csv_path

# 列出CSV文件
csv_files = [file for file in os.listdir(csv_path) if file.endswith('.csv')]

if len(csv_files) == 0:
    logging.error("[ERROR] [list_csv_files] 指定路径下没有找到CSV文件。程序结束。")
    print("指定路径下没有找到CSV文件。程序结束。")
    exit()

logging.info("[INFO] [list_csv_files] 找到以下CSV文件：")
for idx, file in enumerate(csv_files, 1):
    logging.info("[INFO] [list_csv_files] {}. {}".format(idx, file))
    print(f"{idx}. {file}")

# 选择CSV文件
csv_file_index = input("请选择要使用的CSV文件序号：")
csv_file_index = int(csv_file_index) - 1

selected_csv_files = [csv_files[csv_file_index]]
logging.info("[INFO] [select_csv_file] 已选择CSV文件：" + selected_csv_files[0])

# 获取Word文件存储目录
default_word_path = default_folder_path
use_default_word_path = input("是否使用默认的Word文件存储目录？（即代码所在路径）（y/n）：") == "y"

if not use_default_word_path:
    word_path = input("请输入Word文件存储目录：")
    while not os.path.exists(word_path):
        logging.warning("[WARNING] [get_word_path] 路径不存在，请重新输入。")
        word_path = input("请输入Word文件存储目录：")
else:
    word_path = default_word_path

# 列出Word文件
word_files = [file for file in os.listdir(word_path) if file.endswith(('.docx', '.doc'))]

if len(word_files) == 0:
    logging.error("[ERROR] [list_word_files] 指定路径下没有找到Word文件。程序结束。")
    print("指定路径下没有找到Word文件。程序结束。")
    exit()

logging.info("[INFO] [list_word_files] 找到以下Word文件：")
for idx, file in enumerate(word_files, 1):
    logging.info("[INFO] [list_word_files] {}. {}".format(idx, file))
    print(f"{idx}. {file}")

# 选择Word文件
word_file_index = input("请选择要使用的Word文件序号：")
word_file_index = int(word_file_index) - 1

selected_word_files = [word_files[word_file_index]]
logging.info("[INFO] [select_word_file] 已选择Word文件：" + selected_word_files[0])

# 选择CSV文件中的参数作为文件名
with open(os.path.join(csv_path, selected_csv_files[0]), 'r', encoding='utf-8') as csv_file:
    reader = csv.reader(csv_file)
    data = list(reader)

# 提取表头
headers = data[0]
data = data[1:]

logging.info("[INFO] [select_filename_params] 请选择使用CSV文件中的第x个参数作为文件名（默认第一个）：")
for idx, param in enumerate(headers, 1):
    logging.info("[INFO] [select_filename_params] {}. {}".format(idx, param))
    print(f"{idx}. {param}")

filename_param_indices = input("请输入要使用的参数序号（多个参数请使用点“.”分隔）：")
filename_param_indices = [int(idx) - 1 for idx in filename_param_indices.split('.')]

for row in data:
    # 生成文件名
    new_filename = '-'.join([row[idx].strip() for idx in filename_param_indices if row[idx].strip()])
    new_filename = re.sub(filename_regex, '', new_filename)
    logging.debug("[DEBUG] [generate_filename] 生成文件名：" + new_filename)

    # 复制Word文件并保存为新文件名
    for file in selected_word_files:
        doc = docx.Document(os.path.join(word_path, file))
        new_doc = deepcopy(doc)
        new_doc.save(os.path.join(folder_path, f"{new_filename}.docx"))

logging.info("[INFO] [process_completed] 处理完成！")
print("处理完成！")

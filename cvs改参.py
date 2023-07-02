import csv
from datetime import datetime

def convert_date(date_str):
    date = datetime.strptime(date_str, "%Y-%m")
    return date.strftime("%Y年%m月")

def convert_csv(file_path):
    data = []
    with open(file_path, "r") as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader)  # 跳过标题行
        for row in reader:
            date1 = convert_date(row[2])
            date2 = convert_date(row[3])
            row[2] = date1
            row[3] = date2
            data.append(row)

    with open(file_path, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(header)
        writer.writerows(data)


file_path = "output.csv"  # 替换为实际的CSV文件路径
convert_csv(file_path)

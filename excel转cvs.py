import pandas as pd

# 读取Excel文件
excel_file = pd.read_excel('your_file.xlsx')

# 提取相应范围的数据
data = excel_file.iloc[1:42, 1:5]  # 列索引从1到5对应B到E列

# 生成CSV文件
data.to_csv('output.csv', index=False)

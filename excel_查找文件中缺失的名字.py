#实现在一个excel表中查找另外一个excel表不存在的名字
import pandas as pd

# 设置文件路径
file1 = 'zongmingdan.xlsx'  # 第一个 Excel 文件（包含目标工作表）
file2 = 'zuowei.xlsx'  # 第二个 Excel 文件（包含多个工作表）

# 设置第一个文件的工作表和目标列
sheet_name_1 = 'D杰出校友'  # 第一个文件中的工作表名称
name_column_1 = '姓名'   # 第一个文件工作表中要检查的名字列

# 读取第一个文件中的目标工作表
df1 = pd.read_excel(file1, sheet_name=sheet_name_1)

# 获取第一个文件目标工作表中的名字列表
names_file1 = df1[name_column_1].dropna().tolist()

# 加载第二个文件的所有工作表
df2 = pd.read_excel(file2, sheet_name=None)  # sheet_name=None 表示读取所有工作表

# 存储第二个文件中所有工作表的所有值
all_values_file2 = set()

# 遍历第二个文件的所有工作表
for sheet, df in df2.items():
    # 将整个工作表的数据展平，去除空值，存入集合
    all_values_file2.update(df.values.flatten())

# 查找未出现在第二个文件任何工作表中的名字
missing_names = [name for name in names_file1 if name not in all_values_file2]

# 将缺失的名字保存到新的 DataFrame
missing_df = pd.DataFrame(missing_names, columns=[name_column_1])

# 输出缺失的名字到新的 Excel 文件
missing_df.to_excel('missing_names.xlsx', index=False)

print(f"未出现在第二个文件中的名字已保存至 'missing_names.xlsx' 文件")
""""
需求：系统导出的订单文件，一个订单号一行记录，但是一个订单号可能存在多个商品，现在需要将多个商品拆分成多行
分析：将多出的商品单独拿出，根据订单号，商品来提取，然后合并。因为多商品订单里面的商品数量肯定不一样，所以最后合并的肯定会有空白商品。
问题：导出的文件存在列名缺失，从多产品订单的第二个商品开始缺失，列名规则：货品2，货品2数量，货品2成本，以此类推
步骤：
     1.读取文件
     2.补齐缺失的列名 --首先确定有多少列，然后减去已经有列名的，剩下的按规则命名
     3.找出存在多个产品的订单行 --只要是 货品2 这列不为空则是多商品订单
     4.将文件拆分成两部分，一部分是单个商品订单-- part1，另一部分是多个商品订单 -- part2
     5.将多个商品订单（part2）列转行成 part3 --但是有3列转行，分别是货品名，货品数量，货品成本
     6.将 part3 和 part1 合并
"""

import pandas as pd
import os
import time

# file_path = input("请输入文件路径：")
# file_name = input("请输入文件名称：")
time_start = time.time()
# 1.读取文件（输入文件路径和文件名）
f = pd.read_excel(os.path.join(r'C:\Users\HP\Desktop', '副本泛平台利润统计 - 副本.xlsx'), header=11)

# 2.补齐缺失的列名
#   文件有多少列 方法一，f.shape[1] 方法二，f.columns.size
# 2.1 首先获取全部数据的列数
col_num = f.columns.size

# 2.2 然后获取最后一个有列名的列数
# 方法一，df.columns.get_loc('列名')
# 方法二，columns_list = list(df.columns) ，columns_list.index('列名')
col_a = f.columns.get_loc('货品1成本')

# 2.3 两者相减除以3即 t 就是没有列名的商品数，需要按照列名规则命名：货品2，货品2数量，货品2成本，以此类推
t = int((col_num - 1 - col_a)/3)


# 2.4 未命名列获取名字函数
def col_raw_name(b):
    raw_name = 'Unnamed: ' + str(b)
    return raw_name


# 2.5 将未命名列进行重命名
x = 2
while x < t + 1:
    for i in range(col_a + 1, col_num, 3):
        f.rename(columns={col_raw_name(i): "货品"+str(x),
                          col_raw_name(i+1): "货品"+str(x)+"数量",
                          col_raw_name(i+2): "货品"+str(x)+"成本"}, inplace=True)
        x += 1

# 3. 将订单拆分，商品1为 f1,商品2~之后的商品为 f2 ，统一命名为：货品、货品数量、货品成本
# 3.1 注意：只需将多出的商品拆分出来就行
f1 = f.loc[:, :"货品1成本"].rename(columns={"货品1": "货品", "货品1数量": "货品数量", "货品1成本": "货品成本"})
f2 = f[f["货品2"].notna()]
# print(f1.head(10))
# print(f2.head())

# 4. 多商品订单 f2 ,根据订单号、货品、货品数量、货品数量提取这三列数据，然后合并成 f3
f2_merge = []
for x in range(2, t+2):
    cp = f2[["订单号", "货品"+str(x), "货品"+str(x)+"成本", "货品"+str(x)+"数量"]].rename(
        columns={"货品"+str(x): "货品", "货品"+str(x)+"成本": "货品成本", "货品"+str(x)+"数量": "货品数量"})
    f2_merge.append(cp)

f3 = pd.concat(f2_merge)

# 5. 剔除无产品空行
f3.dropna(inplace=True)

# 6. 将 f3 与 f1 合并，根据订单号和付款时间排序，向上填充缺失值，重置索引
result = pd.concat([f1, f3]).sort_values(by=['订单号', '付款日期']).fillna(method='ffill').reset_index(drop=True)

# 7.将结果储存到文件
result.to_excel(os.path.join(r'C:\Users\HP\Desktop', '修改后的文件.xlsx'))

# print(f.columns.values)
time_end = time.time()
print(f'花费时间：{time_end - time_start} s')

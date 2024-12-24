import pandas as pd
from sklearn.preprocessing import StandardScaler
import numpy as np
import openpyxl
from sklearn.cluster import KMeans

# 1.读取数据并完成初步数据清洗
data = pd.read_excel('0-数据整合.xlsx')
cleaned_data = data.dropna()  # 删除含有空白值的行
cleaned_data = cleaned_data[cleaned_data['销售额同比增长率x13'] != '-']  # 删除"销售额同比增长x13"列中值为”-”的数据行
cleaned_data = cleaned_data.iloc[:, :-2]  # 删除最后两列

# 2.数据预处理
# 2.1将非数值型数据转换为数值型数据
cleaned_data['是否是现代终端x6'] = cleaned_data['是否是现代终端x6'].map({'是': 1, '否': 0})
cleaned_data['信用等级x8'] = cleaned_data['信用等级x8'].map({'AAA': 5, 'AA': 4, 'A': 3, 'B': 2, 'C': 1, 'D': 0})
cleaned_data['市场类型x10'] = cleaned_data['市场类型x10'].map({'城网': 1, '农网': 0})
cleaned_data['商圈类型x11'] = cleaned_data['商圈类型x11'].map({
    '办公区': 6,
    '工业区': 4,
    '集贸区': 7,
    '交通枢纽区': 6,
    '居民区': 7,
    '旅游景区': 5,
    '农林渔牧区': 3,
    '其他': 2,
    '商业娱乐区': 8,
    '院校学区': 4
})
cleaned_data['零售业态x12'] = cleaned_data['零售业态x12'].map({
    '便利店': 7,
    '超市': 8,
    '其他': 3,
    '烟草专业店': 9,
    '娱乐服务类': 5
})
cleaned_data['卷烟价格执行情况x15'] = cleaned_data['卷烟价格执行情况x15'].map({
    '很好': 4,
    '较好': 2,
    '一般': 0
})
cleaned_data['配合程度x16'] = cleaned_data['配合程度x16'].map({
    '好': 4,
    '较好': 2,
    '一般': 1
})

# 2.2数据标准化（Z-score标准化）
# 提取要进行标准化处理的列（除第1、2、3、4列外）
columns_to_standardize = cleaned_data.columns[4:]

# 使用Z-score标准化
scaler_zscore = StandardScaler()
data_zscore = scaler_zscore.fit_transform(cleaned_data[columns_to_standardize])

# 将Z-score标准化后的数据转换回DataFrame格式
standardized_data = pd.DataFrame(data_zscore, columns=columns_to_standardize)

# 将标准化后的数据与原数据的前4列合并
cleaned_data = cleaned_data.reset_index(drop=True)  # 重置cleaned_data的索引
final_data = pd.concat([cleaned_data.iloc[:, :4], standardized_data], axis=1)
print(final_data.head())
print('总行数据：', len(final_data))

# 3.根据数据权重计算数据项总和评分
# 定义指标名称和对应的权重字典
weights_dict = {
    '总购进量x1': 0.1979,
    '销售金额x2': 0.1527,
    '一二类烟销量x3': 0.1343,
    '一二类烟金额x4': 0.1336,
    '品牌宽度x5': 0.0465,
    '是否是现代终端x6': 0.0253,
    '电子结算成功率x7': 0.0132,
    '信用等级x8': 0.0034,
    '信用得分x9': 0.0728,
    '市场类型x10': 0.0596,
    '商圈类型x11': 0.0842,
    '零售业态x12': 0.1204,
    '销售额同比增长率x13': 0.0858,
    '卷烟陈列面积x14': 0.0571,
    '卷烟价格执行情况x15': 0.0139,
    '配合程度x16': 0.1405
}

weights_dict_present = {
    '总购进量x1': 0.1979,
    '销售金额x2': 0.1527,
    '一二类烟销量x3': 0.1343,
    '一二类烟金额x4': 0.1336,
    '品牌宽度x5': 0.0465,
    '是否是现代终端x6': 0.0253,
    '电子结算成功率x7': 0.0132,
    '信用等级x8': 0.0034,
    '信用得分x9': 0.0728,
}

weights_dict_potential = {
    '市场类型x10': 0.0596,
    '商圈类型x11': 0.0842,
    '零售业态x12': 0.1204,
    '销售额同比增长率x13': 0.0858,
    '卷烟陈列面积x14': 0.0571,
    '卷烟价格执行情况x15': 0.0139,
    '配合程度x16': 0.1405
}

# 提取用于计算得分的列数据（第5列到第20列）
columns_for_score = final_data.columns[4:20]  # 从第5列到第20列，总和得分
columns_for_score_present = final_data.columns[4:13]  # 从第5列到第13列，当前价值
columns_for_score_potential = final_data.columns[14:20]  # 从第10列到第20列，潜在价值

# 计算每行数据的总和得分
final_data['总和得分'] = 0
final_data['当前价值'] = 0
final_data['潜在价值'] = 0

for col in columns_for_score:
    final_data['总和得分'] += final_data[col] * weights_dict[col]
for col in columns_for_score_present:
    final_data['当前价值'] += final_data[col] * weights_dict_present[col]
for col in columns_for_score_potential:
    final_data['潜在价值'] += final_data[col] * weights_dict_potential[col]

# 对数据按照总和得分进行降序排序
sorted_data = final_data.sort_values(by='总和得分', ascending=False)

# 取前300行数据作为高价值客户
top400_data = sorted_data.head(400)

# 打印最终数据的1,2,3,4和最后一列数据
print(top400_data.iloc[:, [0, 1, 2, 3, -3, -2, -1]].head())

# 4. 用k-means算法对数据进行聚类
# 提取潜在价值、当前价值列的数据作为聚类的特征
X = top400_data[['当前价值']]
Y = top400_data[['潜在价值']]

# 初始化K-Means聚类器，设置聚类数为2
kmeans1 = KMeans(n_clusters=2)
kmeans2 = KMeans(n_clusters=2)

# 对数据进行聚类
kmeans1.fit(X)  # label: 0->好
kmeans2.fit(Y)  # label: 1->好

# 将聚类结果添加到原始数据中作为新的一列
output = top400_data.copy()
output.loc[:, '当前价值聚类'] = kmeans1.labels_
output.loc[:, '潜在价值聚类'] = kmeans2.labels_

print(output.head())

output['当前价值聚类'] = output['当前价值聚类'].map({
    0: '高价值',
    1: '低价值'
})

output['潜在价值聚类'] = output['潜在价值聚类'].map({
    1: '高潜力',
    0: '低潜力'
})

# 将聚类结果写入excel文件
output.to_excel('output.xlsx', index=False)

# 去除指标列
output_info = output.drop(columns_for_score, axis=1)
output_info.to_excel('output_info.xlsx', index=False)

# 筛选潜在价值高的客户
potential_customers = output_info[output_info['潜在价值聚类'] == '高潜力']
print(potential_customers.head())

# 结果写入excel文件
potential_customers.to_excel('潜力客户.xlsx', index=False)

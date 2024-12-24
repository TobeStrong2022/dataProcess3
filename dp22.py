import os
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
import time
import openpyxl

# 1.读取数据并完成初步数据清洗
print("读取数据并完成初步数据清洗...")
data = pd.read_excel('原始数据/客户数据整合.xlsx')
cleaned_data = data.dropna()  # 删除含有空白值的行
cleaned_data = cleaned_data[cleaned_data['销售额同比增长率x13'] != '-']  # 删除"销售额同比增长x13"列中值为”-”的数据行
cleaned_data = cleaned_data.iloc[:, :-2]  # 删除最后两列

# 2.数据预处理
print("数据预处理...")
# 2.1将非数值型数据转换为数值型数据
print("将非数值型数据转换为数值型数据...")
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
print("数据标准化...")
# 提取要进行标准化处理的列（除第1、2、3、4列外）
columns_to_standardize = cleaned_data.columns[4:]

# 使用Z-score标准化
print("使用Z-score标准化...")
scaler_zscore = StandardScaler()
data_zscore = scaler_zscore.fit_transform(cleaned_data[columns_to_standardize])

# 将Z-score标准化后的数据转换回DataFrame格式
standardized_data = pd.DataFrame(data_zscore, columns=columns_to_standardize)

# 将标准化后的数据与原数据的前4列合并
print("数据合并...")
cleaned_data = cleaned_data.reset_index(drop=True)  # 重置cleaned_data的索引
final_data = pd.concat([cleaned_data.iloc[:, :4], standardized_data], axis=1)
# print(final_data.head())
print('总行数据：', len(final_data))

# 3.根据数据权重计算数据项总和评分
print("根据数据权重计算数据项总和评分...")
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

# 取前xxx行数据作为高价值客户, 不取则表示取全部数据
# top400_data = sorted_data.head(400) # 取前400行数据
top400_data = sorted_data  # 取全部数据

# 打印最终数据的1,2,3,4和最后一列数据
# print(top400_data.iloc[:, [0, 1, 2, 3, -3, -2, -1]].head())

# 4. 用k-means算法对数据进行聚类
print("用k-means算法对数据进行聚类...")
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

# print(output.head())

output['当前价值聚类'] = output['当前价值聚类'].map({
    0: '经营状况良好',
    1: '经营状况一般'
})

output['潜在价值聚类'] = output['潜在价值聚类'].map({
    1: '经营潜力佳',
    0: '经营潜力一般'
})

# 将聚类结果写入excel文件
# output.to_excel('output.xlsx', index=False)

# 去除指标列
output_info = output.drop(columns_for_score, axis=1)
# output_info.to_excel('output_info.xlsx', index=False)
customers_guidance = output_info

# 5.生成客户经营指导
print("生成客户经营指导...")
# 根据聚类结果分成的四类客户给出不同的经营指导
# 创建一个新列，用于存放客户经营指导
customers_guidance['客户经营指导'] = '无'
# 若当前价值聚类为高价值，潜在价值聚类为高潜力
customers_guidance.loc[(customers_guidance['当前价值聚类'] == '经营状况良好') & (
            customers_guidance['潜在价值聚类'] == '经营潜力佳'), '客户经营指导'] = \
    '该经客户营状况很好，具有良好的经营基础和较大的发展潜力，应重点维护并助力其进一步提升。建议进一步优化商品组合、强化陈列营销、引导购入更多高价烟品牌，提高客户经营效益。'
# 若当前价值聚类为高价值，潜在价值聚类为低潜力
customers_guidance.loc[(customers_guidance['当前价值聚类'] == '经营状况良好') & (
            customers_guidance['潜在价值聚类'] == '经营潜力一般'), '客户经营指导'] = \
    '该经客户营状况较好，但发展后劲稍显不足，需要巩固现有优势，挖掘潜在增长点。建议深耕现有市场，保持和发挥在总购进量和以及一二类烟销售方面的优势，聚焦现有顾客群体，同时完善规范经营管理，提高经营效益。'
# 若当前价值聚类为低价值，潜在价值聚类为高潜力
customers_guidance.loc[(customers_guidance['当前价值聚类'] == '经营状况一般') & (
            customers_guidance['潜在价值聚类'] == '经营潜力佳'), '客户经营指导'] = \
    '该客户目前经营情况一般，但具有较大的发展潜力，应重点培育。建议加强与客户的沟通，了解客户需求，提供个性化服务，引导客户加强基础建设，根据所处市场类型、商圈特点以及零售业态等特征，选择合适的经营品牌和制定相应的经营策略，提高客户经营效益。'
# 若当前价值聚类为低价值，潜在价值聚类为低潜力
customers_guidance.loc[(customers_guidance['当前价值聚类'] == '经营状况一般') & (
            customers_guidance['潜在价值聚类'] == '经营潜力一般'), '客户经营指导'] = \
    '该客户目前经营情况一般，可能面临一些经营挑战。建议从基础环节入手，逐步改善经营状况，应先聚焦核心产品，集中精力做好几个主打品牌产品，吸引周边固定客源，同时做到规范经营，提高自身信用等级和档位层级，多学习借鉴行业内的成功经验，提高客户经营效益。'

# 生成时间戳
time_stamp = time.strftime('%d%H%M%S', time.localtime(time.time()))

# 结果写入excel文件
customers_guidance.to_excel('分析结果/客户经营指导' + time_stamp + '.xlsx', index=False, sheet_name= '客户经营指导')
# 将运行过程表格写入excel文件的sheet2（使用openpyxl）
wb = openpyxl.load_workbook('分析结果/客户经营指导' + time_stamp + '.xlsx')
# 创建sheet2
ws = wb.create_sheet(title='运行过程')
# 将运行过程表格写入sheet2
for r in dataframe_to_rows(output, index=False, header=True):
    ws.append(r)
# 保存文件
wb.save('分析结果/客户经营指导' + time_stamp + '.xlsx')


print("客户经营指导文件生成成功！")
os.system('start 分析结果/客户经营指导' + time_stamp + '.xlsx')
input("请按任意键退出...")

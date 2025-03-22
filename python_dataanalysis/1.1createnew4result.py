import pandas as pd
import numpy as np

try:
    # 读取原始Excel文件（需确认实际sheet名称）
    df = pd.read_excel('new3根据留存率推算用户增长的公式.xlsx', sheet_name='处理后的日期表')
except Exception as e:
    print(f"文件读取失败: {e}")
    exit()

# 验证必要列是否存在（需确认实际列名）
required_columns = ['序号', '日新增', '天数2', '天数7', '天数30']  # 根据实际列名修改
missing_cols = [col for col in required_columns if col not in df.columns]
if missing_cols:
    print(f"缺失关键列: {', '.join(missing_cols)}")
    exit()

# 创建结果DataFrame
result_df = pd.DataFrame(
    index=pd.Index(range(1, 33), name='Day'),
    columns=['次日加权留存率', '7日加权留存率', '30日加权留存率']
)

# 计算各留存率
for day in result_df.index:
    # ===== 次日留存率计算 =====
    if day == 1:
        result_df.loc[day, '次日加权留存率'] = 0.0
    else:
        # 有效数据范围：序号从1到(day-2) 
        valid_rows = df[df['序号'] <= (day - 1)]  # 修正范围包含所有历史数据
        numerator = valid_rows['天数2'].sum()    # 假设次日留存列名为'次日留存'
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '次日加权留存率'] = numerator / denominator if denominator else np.nan

    # ===== 7日留存率计算 =====
    if day < 7:
        result_df.loc[day, '7日加权留存率'] = 0.0
    else:
        # 有效数据范围：序号从(day-6)到(day-1) 
        valid_rows = df[(df['序号'] >= (day - 6)) & (df['序号'] <= (day - 1))]
        numerator = valid_rows['天数7'].sum()    # 假设7日留存列名为'7日留存'
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '7日加权留存率'] = numerator / denominator if denominator else np.nan

    # ===== 30日留存率计算 =====
    if day < 30:
        result_df.loc[day, '30日加权留存率'] = 0.0
    else:   
        # 有效数据范围：序号从(day-29)到(day-1)
        valid_rows = df[(df['序号'] >= (day - 29)) & (df['序号'] <= (day - 1))]
        numerator = valid_rows['天数30'].sum()   # 假设30日留存列名为'30日留存'
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '30日加权留存率'] = numerator / denominator if denominator else np.nan

# 结果修饰：更美观
    # 结果转置
result_df = result_df.T


# 保存结果
try:
    result_df.to_excel('new4result.xlsx')
    print("文件已成功生成: new4result.xlsx")
    print("数据样例:")
    print(result_df.tail(5))  # 展示最后5天的计算结果
except Exception as e:
    print(f"文件保存失败: {e}")
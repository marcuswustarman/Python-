import pandas as pd
import numpy as np

try:
    # 读取原始Excel文件（确保sheet名称正确）
    df = pd.read_excel('new3根据留存率推算用户增长的公式.xlsx', sheet_name='处理后的日期表')
except Exception as e:
    print(f"文件读取失败: {e}")
    exit()

# 验证必要列是否存在（按实际列名检查）
required_columns = ['序号', '日新增', '天数2', '天数7', '天数30']
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
        # 有效数据范围：序号 <= day-1（累积历史数据）
        valid_rows = df[df['序号'] <= (day - 1)]
        numerator = valid_rows['天数2'].sum()    # 使用实际列名'天数2'
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '次日加权留存率'] = numerator / denominator if denominator else np.nan

    # ===== 7日留存率计算 =====
    if day < 7:
        result_df.loc[day, '7日加权留存率'] = 0.0
    else:
        # 有效数据范围：序号在[day-6, day-1]之间（最近7天窗口）
        valid_rows = df[(df['序号'] >= (day - 6)) & (df['序号'] <= (day - 1))]
        numerator = valid_rows['天数7'].sum()    # 使用实际列名'天数7'
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '7日加权留存率'] = numerator / denominator if denominator else np.nan

    # ===== 30日留存率计算 =====
    if day < 30:
        result_df.loc[day, '30日加权留存率'] = 0.0
    else:
        # 有效数据范围：序号在[day-29, day-1]之间（最近30天窗口）
        valid_rows = df[(df['序号'] >= (day - 29)) & (df['序号'] <= (day - 1))]
        numerator = valid_rows['天数30'].sum()   # 使用实际列名'天数30'
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '30日加权留存率'] = numerator / denominator if denominator else np.nan

# 保存结果
try:
    result_df.to_excel('new4result.xlsx')
    print("文件已生成: new4result.xlsx")
    print("最后5天计算结果预览:")
    print(result_df.tail(5))
except Exception as e:
    print(f"文件保存失败: {e}")
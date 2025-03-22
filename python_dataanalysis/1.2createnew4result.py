import pandas as pd
import numpy as np



try:
    # 读取原始Excel文件
    df = pd.read_excel('new3根据留存率推算用户增长的公式.xlsx', sheet_name='处理后的日期表')
except Exception as e:
    print(f"文件读取失败: {e}")
    exit()

# 验证列是否存在
required_columns = ['序号', '日新增', '天数2', '天数7', '天数30']
missing_cols = [col for col in required_columns if col not in df.columns]
if missing_cols:
    print(f"缺失关键列: {', '.join(missing_cols)}")
    exit()

# 创建结果表（新增“加权平均率”列）
result_df = pd.DataFrame(
    index=pd.Index(range(1, 33), name='Day'),
    columns=['次日加权留存率', '7日加权留存率', '30日加权留存率', '加权平均率']
)

# 计算各列数据
for day in result_df.index:

   
    

    # --- 原有计算逻辑 ---
    # 次日留存率
    if day == 1:
        result_df.loc[day, '次日加权留存率'] = 0.0
    else:
        valid_rows = df[df['序号'] <= day-1]
        numerator = valid_rows['天数2'].sum()
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '次日加权留存率'] = numerator / denominator if denominator else np.nan
    
    # 7日留存率
    if day < 7:
        result_df.loc[day, '7日加权留存率'] = 0.0
    else:
        valid_rows = df[(df['序号'] >= day-6) & (df['序号'] <= day-1)]
        numerator = valid_rows['天数7'].sum()
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '7日加权留存率'] = numerator / denominator if denominator else np.nan
    
    # 30日留存率
    if day < 30:
        result_df.loc[day, '30日加权留存率'] = 0.0
    else:
        valid_rows = df[df['序号'] <= day-1]
        numerator = valid_rows['天数30'].sum()
        denominator = valid_rows['日新增'].sum()
        result_df.loc[day, '30日加权留存率'] = numerator / denominator if denominator else np.nan
    
    # --- 新增加权平均率 ---
    if day == 1:
        result_df.loc[day, '加权平均率'] = 0.0  # 第一天无历史数据

    total = 0
    current_row = 3  # Excel行号（从1开始）
    current_col = day + 4  # Excel列号（从1开始）


    while current_col > 5:
        # 转换为0-based索引
        df_row = current_row + 1
        df_col = current_col - 1
        
        # 检查行列是否有效
        if df_row <= len(df) and df_col <= len(df.columns):
            total += df.iloc[df_row, df_col]
            valid_rows = df[(df['序号'] >= day-6) & (df['序号'] <= day-1)]
            denominator = valid_rows['日新增'].sum() 
           
        current_row -= 1
        current_col += 1


    total = total / denominator
    # 写入结果（假设result_df的索引是天数）
    result_df.loc[day, '加权平均率'] = total



# 数据格式化
result_df = result_df.round(4)  # 保留4位小数

# 保存结果
try:
    result_df.to_excel('new4result.xlsx')
    print("处理成功！输出文件: new4result.xlsx")
    print("数据样例（Day5）:")
    print(result_df.loc[5:5])
except Exception as e:
    print(f"文件保存失败: {e}")
import pandas as pd

def calculate_weighted_retention(input_path, output_path):
    """
    适配当前数据结构的留存率计算程序
    输入文件结构：
    - 第0列：日期/历史标识
    - 第1列：序号
    - 第2列：日新增
    - 第3列：天数/日活标识
    - 第4-34列：留存数据（对应Day1-Day30）
    """
    try:
        # 读取数据，跳过第一行（原示例中的无效标题）``
        df = pd.read_excel(input_path, header=None, skiprows=1)
        print("\n数据读取成功，数据结构预览：")
        print(df.head(3))
    except Exception as e:
        print(f"\n文件读取失败：{str(e)}")
        return

    # 数据清洗
    # 筛选有效数据行（过滤空行和说明行）
    valid_data = df[df[1].apply(lambda x: str(x).isdigit())].copy()
    
    # 列索引映射
    COLUMN_MAPPING = {
        'date_col': 0,        # 日期列
        'new_users_col': 2,   # 日新增列
        'retention_start': 4  # 留存数据起始列（Day1）
    }

    # 初始化结果表
    result = pd.DataFrame(
        index=['次日留存率', '7日留存率', '30日留存率'],
        columns=[f'Day{i}' for i in range(1, 31)]
    )

    # 遍历每一天的新增数据
    for idx, row in valid_data.iterrows():
        day_num = int(row[1])  # 获取序号列的日期编号
        
        try:
            # 获取关键数据
            new_users = row[COLUMN_MAPPING['new_users_col']]
            
            # 计算留存率（按列偏移计算）
            retention_rates = {
                '次日留存率': (row[COLUMN_MAPPING['retention_start']] / new_users) * 100,
                '7日留存率': (row[COLUMN_MAPPING['retention_start']+6] / new_users) * 100,
                '30日留存率': (row[COLUMN_MAPPING['retention_start']+29] / new_users) * 100
            }
            
            # 填入结果表
            for metric, rate in retention_rates.items():
                result.loc[metric, f'Day{day_num}'] = f"{rate:.2f}%"
                
        except Exception as e:
            print(f"计算Day{day_num}时出错：{str(e)}")
            continue

    # 导出结果
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        result.to_excel(writer, sheet_name='留存率报告')
        
        # 设置格式
        workbook = writer.book
        worksheet = writer.sheets['留存率报告']
        
        # 设置列宽
        worksheet.set_column('A:AE', 15)
        
        # 设置数字格式
        percent_format = workbook.add_format({'num_format': '0.00%'})
        for col in range(1, 31):
            worksheet.set_column(col, col, None, percent_format)

    print(f"\n处理完成，结果已保存至：{output_path}")

# 使用示例
if __name__ == "__main__":
    input_file = 'new2根据留存率推算用户增长的公式.xlsx'  # 根据实际文件名修改
    output_file = '加权留存率报告.xlsx'
    
    calculate_weighted_retention(
        input_path=input_file,
        output_path=output_file
    )
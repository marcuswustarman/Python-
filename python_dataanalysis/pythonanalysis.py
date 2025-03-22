
import pandas as pd
import numpy as np

# 读取Excel数据（需提前安装openpyxl库）
df = pd.read_excel("retention_data.xlsx", sheet_name="Sheet1")

# 定义需要计算的留存天数范围（示例计算3日加权留存）
retention_days = 3  # 可根据实际需要修改
weight = [0.5, 0.3, 0.2]  # 自定义权重（需满足 sum=1）

# 预处理步骤
def calculate_weighted_retention(row):
    """计算单行的加权留存率"""
    # 提取留存率数值并转换为小数
    retention_values = [
        float(str(x).replace("%", "")) / 100 
        for x in row[[f"Day{i}_Retention" for i in range(1, retention_days+1)]]
    ]
    # 计算加权值
    return np.dot(retention_values, weight)

# 应用计算
df["Weighted_Retention"] = df.apply(calculate_weighted_retention, axis=1)

# 结果保存
df.to_excel("weighted_retention_result.xlsx", index=False)

# 验证权重总和为1
assert abs(sum(weight) - 1.0) < 1e-6, "权重总和必须等于1"
print(df)
# 显示前3行结果示例
print(df[["日期", "Weighted_Retention"]].head(3))
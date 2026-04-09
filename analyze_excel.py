import pandas as pd
import os

excel_path = r"temp\ISO17025(2017)實驗室管理系統文件總覽表.xlsx"

# 嘗試讀取，不指定 header 看看原始樣貌
df = pd.read_excel(excel_path, header=None)
print("Shape:", df.shape)
print("First 15 rows, all columns:")
print(df.head(15).to_string())

# 尋找可能包含 "測試方法" 或 "檔案編號" 的列
for index, row in df.iterrows():
    row_str = " ".join(map(str, row.values))
    if "測試方法" in row_str or "檔案編號" in row_str:
        print(f"Found keyword at row {index}: {row_str}")

# 保存為 csv 以便分析
df.to_csv("full_excel_analysis.csv", index=False, encoding="utf-8-sig")

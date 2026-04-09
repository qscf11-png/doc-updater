import pandas as pd
import os

excel_path = r"C:\Users\TK_Tsai\OneDrive - Moxa Inc\桌面\文件便更\ISO17025(2017)實驗室管理系統文件總覽表.xlsx"

if not os.path.exists(excel_path):
    print(f"File not found: {excel_path}")
else:
    df = pd.read_excel(excel_path)
    print("Columns:", df.columns.tolist())
    # 根據使用者描述：F 欄是舊的，B 欄是新的
    # 注意 pandas 0-indexed: B 是 1, F 是 5
    # 列出前 10 筆資料以便確認邏輯
    subset = df.iloc[:, [1, 5]].dropna().head(20)
    print("--- Mapping Sample (Column B -> Column F) ---")
    print(subset)
    
    # 儲存到一個暫存檔以便我也能看到更多
    subset_all = df.iloc[:, [1, 5]].dropna()
    subset_all.to_csv("mapping.csv", index=False, encoding="utf-8-sig")


import pandas as pd
import openpyxl

# CSVファイルを読み込む
csv_data = pd.read_csv("Financial_Forecast_Revision_JA.csv")

# Excelファイルを開く
wb = openpyxl.load_workbook("ソフトバンク調査.xlsx")
sheet = wb["決算説明会"]

# C2セルに期初予想の括弧内の値を転記
if "（" in csv_data.iloc[0, 1] and "）" in csv_data.iloc[0, 1]:
    initial_forecast_date = csv_data.iloc[0, 1].split("（")[1].split("）")[0]
    sheet["C2"] = initial_forecast_date
else:
    sheet["C2"] = ""

# 売上高の値を転記
sheet["B3"] = csv_data.iloc[0, 1]
sheet["D3"] = csv_data.iloc[0, 2]
sheet["E3"] = csv_data.iloc[0, 3]

# 営業利益の値を転記
sheet["B4"] = csv_data.iloc[1, 1]
sheet["D4"] = csv_data.iloc[1, 2]
sheet["E4"] = csv_data.iloc[1, 3]

# 親会社の所有者に帰属する純利益の値を転記
sheet["B5"] = csv_data.iloc[2, 1]
sheet["D5"] = csv_data.iloc[2, 2]
sheet["E5"] = csv_data.iloc[2, 3]

# Excelファイルを保存
wb.save("ソフトバンク調査.xlsx")
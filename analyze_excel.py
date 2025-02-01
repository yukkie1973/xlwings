import xlwings as xw             # 1. xlwingsをインポート
import sys

# 第1引数を取得
args = sys.argv
if len(args) < 2:
    print("Please input Excel file path.")
    sys.exit()
excel_file = args[1]

wb = xw.Book(excel_file) # 2. ブックを開く

# VBAプロジェクトを取得
vba_project = wb.api.VBProject

# VBAマクロの名前をリストアップ
print("List of VBA Macros:")
for component in vba_project.VBComponents:
    if component.Type == 1:  # 1 = Standard Module
        code_module = component.CodeModule
        for i in range(1, code_module.CountOfLines + 1):
            line = code_module.Lines(i, 1)
            if line.strip().startswith("Sub "):
                sub_name = line.strip().split(" ")[1].split("(")[0]
                print(f"  Sub Name: {sub_name}")
            if line.strip().startswith("Function "):
                func_name = line.strip().split(" ")[1].split("(")[0]
                print(f"  Function Name: {func_name}")

# シート名をリストアップ
print("List of Sheet Names:")
for sheet in wb.sheets:
    print(sheet.name)
    # シート内にグラフがあればグラフ名もリストアップ
    for chart in sheet.charts:
        print(f"  グラフ : {chart.name}")
        try:
            # グラフの元データが存在するシートを取得
            #print(f"    グラフの元データ : {chart.values_only}")
            # formula = chart.api.SeriesCollection(1).Formula
            #formula = chart.api.SeriesCollection(1).Formula
            #print(f"    グラフの元データformula : {formula}")
            #data_sheet_name = formula.split('!')[0].replace("'", "")
            #print(f"    グラフの元データ : {data_sheet_name}")
            # グラフのデータ範囲を取得
            #print(f"    データ範囲 : {formula}")
            # グラフの種類を取得
            print(f"    グラフの種類 : {chart.chart_type}")
        except AttributeError as e:
            print(f"    Error retrieving chart data: {e}")

    # シート内にテーブルがあればテーブル名もリストアップ
    for table in sheet.tables:
        print(f"  テーブル : {table.name}")
    # シート内にピボットテーブルがあればピボットテーブル名もリストアップ
    for pivot_table in sheet.api.PivotTables():
        print(f"  ピボットテーブル : {pivot_table.Name}")


# VBAマクロをテキストファイル出力
with open('vba_macros.txt', 'w') as f:
    for component in vba_project.VBComponents:
        if component.Type == 1:  # 1 = Standard Module
            code_module = component.CodeModule
            for i in range(1, code_module.CountOfLines + 1):
                line = code_module.Lines(i, 1) + '\n'
                f.write(line)

# Excelファイルを閉じる
wb.close()

# マクロのSub名一覧を取得
#print(wb.macro('SubName').module)

# macro()                          # 4. マクロを実行
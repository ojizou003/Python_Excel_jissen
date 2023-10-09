'''自動で日報を作成'''

import datetime

import openpyxl

# フォルダとファイルのPATHを記述して、変数に格納する
excel_file_path = 'excel/'
file_name = 'A社_日報雛形.xlsx'

# 変数read_wbに、読み込んだExcelファイルを格納する
read_wb = openpyxl.load_workbook(excel_file_path + file_name)

# 使用するシート(最初の1枚)を選択する
read_ws = read_wb.worksheets[0]

# 変数today_dateに今日の日付を格納する
today_date = datetime.date.today()

# 今日の日付を格納する
read_ws['C3'] = today_date

# その他の基本情報を、該当するセルに記入する
read_ws['C4'] = '9:00'
read_ws['C5'] = '18:00'
read_ws['E3'] = '経理部'
read_ws['E4'] = 'おぢぞう'

# シート名を今日の日付に変更する
read_ws.title = today_date.strftime("%Y-%m-%d")

def write_plans(plans,is_today = False):
    if is_today:
        target_rows = read_ws['b13:b17']
        template = ['日報作成']
    else:
        target_rows = read_ws['b7:b11']
        template = ['A社定例MTG', 'チーム内会議(木曜日)']

    plans = template + plans

    for plan,empty_cell in zip(plans,target_rows):
            empty_cell[0].value = plan

# 関数write_plans()を使って、タスクの記入をおこなう
write_plans(['B社PPT資料作成'])
write_plans(['B社用資料の骨子作成','部下の指導','同僚とランチ'],is_today=True)

# 変更結果を保存する
output_file_name = f'日報_おぢぞう_{today_date}.xlsx'
read_wb.save(excel_file_path + output_file_name)
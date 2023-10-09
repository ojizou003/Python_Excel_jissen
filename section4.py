'''パスワードを自動生成し、ブック・シートを保護'''

import string
import secrets

import openpyxl

# パスワードを生成するコードを関数化する
def get_new_password(password_length = 10,is_symbol = False):
    alphabet = string.ascii_letters + string.digits

    if is_symbol:
        alphabet += string.punctuation

    password = ''
    for _ in range(password_length):
        password += secrets.choice(alphabet)

    print(password)
    return(password)

# フォルダとファイルのPATHを記述して、変数に格納する
excel_file_path ='excel/'
file_name = 'Book1.xlsx'

# 変数read_wbに、読み込んだExcelファイルを格納する
read_wb = openpyxl.load_workbook(excel_file_path + file_name)

# ブックの保護をかける準備をする
read_wb.security = openpyxl.workbook.protection.WorkbookProtection()

# パラメーターの設定をおこなう
read_wb.security.lockStructure = True
read_wb.security.workbookPassword = get_new_password(10,1)

# v2で変更結果を保存する
read_wb.save(excel_file_path + 'v2_' + file_name)

# 保護するシート(最初の1枚)を選択する
read_ws = read_wb.worksheets[0]

# シート内のセル保護を有効にする
read_ws.protection.sheet  = True

# シート内セルをパスワード保護する
read_ws.protection.password = get_new_password(10)

# v3で変更結果を保存する
read_wb.save(excel_file_path + 'v3_' + file_name)
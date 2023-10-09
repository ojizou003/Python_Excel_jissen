# 【Python x Excel 実践編】 はやたす youtube

2023/10/8 ~ 10/9  

## 他ファイルの内容を「自動でコピペ」する方法  

anaconda プロンプトでPythonを実行  
```python *.py```  

Google Colaboratoryのノートブックファイルをパイソンファイルに変換  
``ファイル --> ダウンロード --> .py をダウンロード``

Pythonを使ってフォルダ・ファイルの一覧を確認できるようにしたいので、globをインポート  
```from glob import glob```  

ワークシートをクリアにするためには、もう一度新しい変数wsを作成しなおす  

## コメント一覧を「自動」で取得する方法

Excel操作 -- Ctrl＋Shift＋Lでソート  

## パスワードも自動生成！ブック・シートをパスワードで保護しよう

Pythonファイルを実行したら、Excelファイルで使うパスワードの生成から設定まで、自動でおこなう  

import string  
import secrets  
import openpyxl  

パスワードを生成する手順

1. パスワードで使う文字列を宣言しておく  
    ```alphabet = string.ascii_letters + string.```  
2. 宣言しておいた文字列から好きな文字数のパスワードを生成する  

    ```python
    s = ''
    for _ in range(20):
        s += secrets.choice(alphabet)
    ```

string.punctuation ..記号

Excelブックの保護状況を確認する  
```read_wb.security```  

シートのパスワード保護状況を確認する  
```read_ws.protection.sheet```  

## 自動で日報を作成しよう

結合されているセルの情報を取得する  
```read_ws.merged_cells.ranges```  

シート名を今日の日付に変更する  
```read_ws.title = today_date.strftime('%y-%m-%d')```  
※シート名は文字列である必要がある  

今週のタスクは3つ、空欄は5つのように、個数はバラバラだけど空欄の先頭3つだけ値を埋めていきたいとき、Python文法のzip()を使ってあげれば解決できる  

＜補足＞  
関数の中で変更された変数の中身は、必ずしも関数の外で反映されないわけではない

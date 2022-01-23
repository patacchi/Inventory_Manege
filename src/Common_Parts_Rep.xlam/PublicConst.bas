Attribute VB_Name = "PublicConst"
Option Explicit
'データベース全般の定数定義
'DBファイルの拡張子
Public Const DB_FILE_EXETENSION_ACCDB As String = "accdb"       'DBファイルの拡張子その１(lcaseにして）
Public Const DB_FILE_EXETENSION_XLAM = "xlam"                   'エクセルアドオンファイル
Public Const DB_FILE_EXETENSION_XLSM = "xlsm"                   'エクセルマクロ有効ブック
Public Const DB_FILE_EXETENSION_XLSX = "xlsx"                   'エクセルXLS形式ブック
Public Const DB_FILE_EXETENSION_XLSB = "xlsb"                   'エクセルバイナリ形式ブック
Public Const DB_FILE_EXETENSION_XLS = "xls"                     'Excel97までのエクセルブック形式
Public Const DB_FILE_EXETENSION_CSV = "csv"                     'CSVファイル形式
'DBファイルの列挙体
Public Enum DB_file_exetension
    accdb_dext = 1
    xlam_dext = 2
    xlsm_dext = 3
    xlsx_dext = 4
    xlsb_dext = 5
    xls_dext = 6
    csv_dext = 7
End Enum
'各テーブルにはDefaultとしてInputDateが入る
Public Const INPUT_DATE As String = "InputDate"
'Tempデータベース、Excelファイルは一時テーブルに格納した方が上手くいくみたいなので、とりあえず一時テーブルのみを置くデータベースファイル
Public Const TEMP_DB_FILENAME As String = "DB_Temp_Local.accdb"       '全DB共通
'-------リスト表示のための定数定義
'MS ゴシック（等幅）文字サイズ9ptの場合
Public Const sglChrLengthToPoint = 4.1
Public Const longMINIMULPOINT = 50
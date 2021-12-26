Attribute VB_Name = "PublicConst"
Option Explicit
'データベース全般の定数定義
'DBファイルの拡張子
Public Const DB_FILE_EXETENSION1 As String = "accdb"            'DBファイルの拡張子その１(lcaseにして）
'DBファイルの列挙体
Public Enum DB_file_exetension
    accdb_dext = 0
End Enum
'-------リスト表示のための定数定義
'MS ゴシック（等幅）文字サイズ9ptの場合
Public Const sglChrLengthToPoint = 4.1
Public Const longMINIMULPOINT = 50
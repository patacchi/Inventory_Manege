Attribute VB_Name = "PublicConst"
Option Explicit
'MsformsでMouseイベントのButton定数値
Public Const vbMouseLeft As Integer = 1                         '左クリック時のコード
Public Const vbMouseRight As Integer = 2                        '右クリック時のコード
Public Const vbMouseMiddle As Integer = 4                       '中央ボタンクリック時のコード
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
'フィールド抽出条件文字列(SQL)
Public Const SQL_F_EQUAL As String = " = "                      'SQLのWHERE条件式などでフィールド間を連結する文字列 =
Public Const SQL_F_NOT_EQUAL As String = " <> "                 '<>
Public Const SQL_F_TRIM_PREFIX As String = "TRIM("              'TRIM$使う時にフィールドの前に付くプレフィックス
Public Const SQL_ISNUMERIC_0FieldName As String = "IIF(ISNUMERIC({0}),CDBL({0}),0)"   '0にフィールド名を入れる、数値はDoubleで入力するので、型変換をする、これをしないとデータ型エラーが出る
Public Const SQL_ISNULLTRIM_0FieldName As String = "IIF(ISNULL(TRIM({0})),"""",TRIM({0}))"          'Trimの結果がNullだった場合は固定の文字列を入れる
Public Const SQL_F_TRIM_SUFFIX As String = ")"                  'TRIM使用時のサフィックス
Public Const SQL_F_CONNECT_OR As String = " OR "                '条件が複数時において、次の条件との間に挟む語句 OR
Public Const SQL_F_CONNECT_AND As String = " AND "              'AND
Public Const SQL_F_CONNECT_COMMA As String = ","                ', これだけは前後にスペース入れない方が良い
'抽出条件文字列Enum
Public Enum Enum_SQL_F_Condition
    Equal_sfc = 1
    NOT_Equal_sfc = 2
    Trim_Prefix_sfc = 3
    Trim_Suffix_sfc = 4
    Connect_OR_sfc = 5
    Connect_AND_sfc = 6
    Connect_Comma_sfc = 7
End Enum
'各テーブルにはDefaultとしてInputDateが入る
Public Const INPUT_DATE As String = "F_InputDate"
'Tempデータベース、Excelファイルは一時テーブルに格納した方が上手くいくみたいなので、とりあえず一時テーブルのみを置くデータベースファイル
Public Const TEMP_DB_FILENAME As String = "DB_Temp_Local.accdb"       '全DB共通
'-------リスト表示のための定数定義
'MS ゴシック（等幅）文字サイズ9ptの場合
Public Const sglChrLengthToPoint = 4.1
Public Const longMINIMULPOINT = 50
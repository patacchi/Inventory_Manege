Attribute VB_Name = "CAT_CONST"
Option Explicit
'''CATテーブルに関する定数を置いておく
'Global
Public Const CAT_DB_FILENAME As String = "CAT_Find.accdb"                   'CATコードのDBファイル名
'各テーブルにはDefaultとしてInputDateが入る
Public Const INPUT_DATE As String = "InputDate"
'Header
Public Const T_CAT_HEADER As String = "T_CAT_Header"                        'CATコードのヘッダ情報を詰め込んだテーブル（起点）
Public Const F_CAT_HEADER As String = "F_CAT_Header"                        'CATヘッダのフィールド名、これが通常機種名になる
Public Const F_CAT_DESCRIPTIONTABLE As String = "F_CAT_DescriptionTable"    'Descriptionテーブルのフィールド名
Public Const F_CAT_DETAILTABLE As String = "F_CAT_DetailTable"              '詳細（仕様）テーブル名のフィールド名
Public Const F_CAT_SPECIALTABLE As String = "F_CAT_SpecialTable"            '特殊条件を記載したテーブル名のフィールド名
'CATのヘッダテーブル定義のEnum、実際の値はclsEnumで定義する
'メンバー名重複防止すると頭にEnum識別名をつけないとコンパイルエラーになるため、メンバー名は重複しないようにする
'重複すると 名前が適切ではありません のエラーが発生する
'重複防止のため、サフィックスとして原則、_??? を付加する
'マスターテーブルのサフィックスは_?m?? とする
Public Enum Enum_CAT_Header
    T_Name_chd = 0              'Header Tableそのものの名前
    F_Header_chd = 1            '各CATコードのヘッダ部分
    F_DescriptionTable_chd = 2  'Description定義のテーブル名
    F_DetailTable_chd = 3       'Detailのテーブル名
    F_SpecialTable_chd = 4      'Specialのテーブル名
    F_InputDate_chd = 5         '入力日時
End Enum
'（各機種）_Description
'各桁の概要（種類）を格納する テーブル名は T_CAT_（各機種名）_Description とする
'説明自体は外部キーとして設定し、親テーブルはT_CAT_M_Descriptionとする
'数機種で共有する前提で設計する
Public Const T_CAT_DESCRIPTION_0kishu As String = "T_CAT_{0}_Description"   'Detailテーブル名 {0}に各機種を埋め込む
Public Const F_CAT_DIGIT_ROW As String = "F_CAT_Digit_Row"                  '桁数のビットのフラグを立てたLongの数、n桁目は2^(n-1) 3桁目なら4、各テーブルにはこの値をセットする
Public Const F_CAT_DESCRIPTION_ID As String = "F_CAT_Description_ID"        'DescriptionのIDが入る
'DescriptionテーブルのEnum
Public Enum Enum_CAT_Description
    T_Name_0_Kishu_cdc = 0          'インデックス0（機種名に置換が必要）Descriptionテーブル名
    F_Digit_Row_cdc = 1             'フラグ立てたLongの数
    F_Descriptoin_ID_cdc = 2        'DescriptionのID、実態はマスターテーブル参照
    F_InputDate_cdc = 3             '入力日時
End Enum
'（各機種）_Detail
'CATコードの各桁位置に対する説明を格納する（メインテーブル）
'テーブル名は T_CAT_（各機種名）_Detailとする 一つのテーブルに複数機種を格納する可能性がある
Public Const T_CAT_DETAIL_0kishu As String = "T_CAT_{0}_Detail"             '{0}に機種名を埋め込む
Public Const F_CAT_CHR As String = "F_CAT_Chr"                              'その桁に入る文字列
Public Const F_CAT_DETAIL_ID As String = "F_CAT_Detail_ID"                  'DetailのIDが入る
'DetailテーブルのEnum
Public Enum Enum_CAT_Detail
    T_Name_0_Kishu_cdt = 0          'Detaiテーブル名 0 （機種名置換）が必要
    F_Digit_Row_cdt = 1             '桁数フラグのLong
    F_Chr_cdt = 2                   '桁に入る文字
    F_Detail_ID_cdt = 3             'DetailのID
    F_InputDate_cdt = 4             '入力日時
End Enum
'（各機種）_Special
'特殊な組み合わせで表現が変わるものを集めたテーブル
'条件、結果（改変実行）どちらもJSONで格納する
Public Const T_CAT_SPECIAL_0kishu As String = "T_CAT_{0}_Special"           '{0}に機種名が入る
Public Const F_CAT_CONDITION As String = "F_CAT_Condition"                  '条件をJSONで格納
Public Const F_CAT_EXECUTE As String = "F_CAT_Excute"                       '改変する内容をJSONで格納
'SpecialテーブルのEnum
Public Enum Enum_CAT_Special
    T_Name_0_Kishu_csp = 0          'Specialのテーブル名、{0}を機種名置換必要
    F_Condition_csp = 1             '条件フィールド
    F_Execute_csp = 2               '改変内容フィールド
    F_InputDate_csp = 3             '入力日時
End Enum
'Descriptionマスターテーブル
'実際の概要はこっちに入る。親テーブル
Public Const T_CAT_Description_MASTER As String = "T_CAT_M_Description"     'Descriptionのマスター
Public Const F_CAT_DESCRIPTION_TEXT As String = "F_CAT_Description"         '実際の概要の内容が入るフィールド
'DescriptionマスターテーブルのEnum
Public Enum Enum_CAT_M_Description
    T_Name_cmdc = 0                  'Descriptionマスターのテーブル名
    F_Description_ID_cmdc = 1        'Descriptionテーブルと共用、こっちが親
    F_Digit_Row_cmdc = 2             'Descriptionテーブルと共用
    F_Description_Text_cmdc = 3      'Descriptionの本体
    F_InputDate_cmdc = 4             '入力日時
End Enum
'Detailマスターテーブル
Public Const T_CAT_DETAIL_MASTER As String = "T_CAT_M_Detail"               'Detailマスターのテーブル名
Public Const F_CAT_DETAIL_TEXT As String = "F_CAT_Detail"                   '実際の仕様詳細が入るフィールド
'DetailマスターテーブルのEnum
Public Enum Enum_CAT_M_Detail
    T_Name_cmdt = 0                  'Detailマスターのテーブル名
    F_Detail_ID_cmdt = 1             'Detailテーブルと共用、こっちが親
    F_Digit_Row_cmdt = 2             'Detailテーブルと共用
    F_Chr_cmdt = 3                   'Detailテーブルと共用
    F_Detail_Text_cmdt = 4           'Detailの本体
    F_InputDate_cmdt = 5             '入力日時
End Enum
'桁数マスターテーブル
Public Const T_CAT_DIGIT_MASTER As String = "T_M_Digit"                 '桁数とLongの数の対応テーブル
Public Const F_CAT_DIGIT_OFFSET As String = "F_CAT_DigitOffset"             'ヘッダの文字の最後を0とした文字位置のオフセット、将来JSON(String配列）で格納するかも？
'桁数マスターテーブルのEnum
Public Enum Enum_CAT_M_Digit
    T_Name_cmdg = 0                 '桁数マスターテーブルの名前
    F_Digit_Offset_cmdg = 1         'ヘッダの文字の最後を0とした文字位置のオフセット（1文字目は1から始まる）
    F_Digit_Row_cmdg = 2            '他のテーブルと共用、こちらが親
End Enum
'桁数フィールド変換用（経過措置）、全部終わったら削除する
Public Const F_DIGIT_UPDATE As String = "F_Digit_Update"                    '桁数形式変換が完了したかどうか、完了したらTrueをセットする
'一時利用フィールドEnum
Public Enum CAT_Tmp
    F_Digit_Update_ctm = 0
End Enum
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'SQL定義
'既存機種用に機種名一覧を取り出すSQL
Public Const SQL_KISHU_LIST As String = "SELECT " & F_CAT_HEADER & "," & F_CAT_DESCRIPTIONTABLE & "," & F_CAT_DETAILTABLE & "," & F_CAT_SPECIALTABLE & vbCrLf _
                                        & " FROM " & T_CAT_HEADER
'フィールド追加SQL
Public Const SQL_APPEND_FIELD_0Tableneme_1fieldname_2DataType As String = "ALTER TABLE {0} ADD COLUMN {1} {2}"      '{0}にTableNameを{1}にフィールド名を {2}にフィールドタイプを入れる
'フィールドデータ型Enum
Public Enum ACCDB_Data_Type
    'LongText
    'データサイズを指定(255以下)するとShortTextになる
    'データサイズを指定しないLongTextだと色々制限があるので非推奨
    Text_typ = 1
    'ShortText
    'あまり使わないほうがいいかも
    Char_ShortText_typ = 2
    Integer_typ = 3
    BIT_typ = 4
    Boolean_typ = 4
    COUNTER_typ = 5
    AUTOINCREMENT_typ = 5
    Decimal_typ = 6
    Single_typ = 7
    Double_Typ = 8
End Enum
'フィールド削除SQL
Public Const SQL_DELETE_FIELD_0Tablename_1Fieldname As String = "ALTER TABLE {0} DROP COLUMN {1}"                   '{0}にTableNameを、{1}にフィールド名を入れる
'フィールドデータ型変更SQL
Public Const SQL_CHANGE_DATATYPE_0Tablename_1Fieldname_2DataType As String = "ALTER TABLE {0} ALTER COLUMN {1} {2}" '0にTablename、1にFieldName、2にDataTypeを入れる
'InputDate .fff修正SQL UPDATE でテーブルに INNER JOIN を使用
'{0}にTableNameを入れる
Public Const SQL_FIX_INPUTDATE_0_TableName As String = "UPDATE " & vbCrLf & _
    "{0} AS Torigin" & vbCrLf & _
    "   INNER JOIN" & vbCrLf & _
    "       (" & vbCrLf & _
    "       SELECT InputDate," & vbCrLf & _
    "       REPLACE(InputDate, ""fff"", ""000"") As InputDate_Replace" & vbCrLf & _
    "       FROM {0}" & vbCrLf & _
    "       ) AS T1" & vbCrLf & _
    "   ON Torigin.InputDate = T1.InputDate" & vbCrLf & _
    "SET Torigin.InputDate = T1.InputDate_Replace" & vbCrLf & _
    "WHERE Torigin.InputDate <> T1.InputDate_Replace"
'DigitOffset → DigitRow
'{0}にTableName、{1}にDigitOffset {2}にDigitRow {3} にDigitUpdateが入る
Public Const SQL_FIX_DIGITOFFSET_0_TableName_1_DigitOffset_2_DigitRow_3_DigitUpdate As String = "UPDATE" & vbCrLf & _
"   {0} As Torigin" & vbCrLf & _
"   INNER JOIN T_M_Digit as T1" & vbCrLf & _
"   ON Torigin.{1} = T1.{1}" & vbCrLf & _
"SET" & vbCrLf & _
"   Torigin.{2} = T1.{2} ," & vbCrLf & _
"   Torigin.{3} = True" & vbCrLf & _
"WHERE" & vbCrLf & _
"   Torigin.{2} <> T1.{2}" & vbCrLf & _
"   OR Torigin.{2} IS NULL"
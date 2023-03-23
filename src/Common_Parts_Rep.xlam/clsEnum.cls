VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''Enumの値から文字列情報を取り出すためのクラス
'''Enum毎にDictionaryを作り、Enumの値をキーとして、直接文字列を埋め込む（できれば他でConstで定義する）
'''
Option Explicit
Private dicDBFileExetension As Dictionary
Private dicAccdbDataType As Dictionary
Private dicCATHeaderField As Dictionary
Private dicCATDescriptionField As Dictionary
Private dicCATDetailField As Dictionary
Private dicCATSpecialField As Dictionary
Private dicCATMasterDescriptionField As Dictionary
Private dicCATMasterDetailField As Dictionary
Private dicMasterDigitField As Dictionary
Private dicCATTmpField As Dictionary
Private dicINVMasterParts As Dictionary
Private dicINVMasterTana As Dictionary
Private dicSHZaikoParts As Dictionary
Private dicSHZaikoKeyPartsValueEnum As Dictionary               '在庫情報シートで、Keyがフィールド名、ValueがEnumに逆転してるDictionary
Private dicSQLFieldCondition As Dictionary
Private dicSQLINVAlias As Dictionary
Private dicCSVTana As Dictionary
Private dicSHZaikoNeedTrim As Dictionary
Private dic_ReverseCSVTana As Dictionary                         '棚卸CSVでkeyとValueが反転したDictonary
Private dicCommonTanaCSV As Dictionary                          '棚卸CSVでDBと共通のフィールドのEnumを格納したDictionary
Private dicCSVTanaNeedTrim As Dictionary                        '棚卸CSVでTrimが必要なフィールドのEnumを格納したDictionary
Private dicSHReceive As Dictionary                              '受入れ情報シートのフィールドEnum定義を格納したDictionary
'コンストラクタ
Private Sub Class_Initialize()
    Call Initial_Valueset
End Sub
Private Sub Initial_Valueset()
    'クラス初期化時にEnum毎のDictionaryに値を入れる
''DBファイルの列挙体
'Public Enum DB_file_exetension
'    accdb_dext = 1
'    xlam_dext = 2
'    xlsm_dext = 3
'    xlsx_dext = 4
'    xlsb_dext = 5
'    csv_dext = 6
'End Enum
    If dicDBFileExetension Is Nothing Then
        Set dicDBFileExetension = New Dictionary
    End If
    dicDBFileExetension.Add DB_file_exetension.accdb_dext, PublicConst.DB_FILE_EXETENSION_ACCDB
    dicDBFileExetension.Add DB_file_exetension.xlam_dext, PublicConst.DB_FILE_EXETENSION_XLAM
    dicDBFileExetension.Add DB_file_exetension.xlsm_dext, PublicConst.DB_FILE_EXETENSION_XLSM
    dicDBFileExetension.Add DB_file_exetension.xlsx_dext, PublicConst.DB_FILE_EXETENSION_XLSX
    dicDBFileExetension.Add DB_file_exetension.xlsb_dext, PublicConst.DB_FILE_EXETENSION_XLSB
    dicDBFileExetension.Add DB_file_exetension.xls_dext, PublicConst.DB_FILE_EXETENSION_XLS
    dicDBFileExetension.Add DB_file_exetension.csv_dext, PublicConst.DB_FILE_EXETENSION_CSV
''抽出条件文字列Enum
'Public Enum_SQL_F_Condition
'    Equal_sfc = 1
'    NOT_Equal_sfc = 2
'    Trim_Prefix_sfc = 3
'    Trim_Suffix_sfc = 4
'    Connect_OR_sfc = 5
'    Connect_AND_sfc = 6
'    Connect_Comma_sfc = 7
'End Enum
    If dicSQLFieldCondition Is Nothing Then
        Set dicSQLFieldCondition = New Dictionary
    End If
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.Equal_sfc, SQL_F_EQUAL
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.NOT_Equal_sfc, SQL_F_NOT_EQUAL
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.Trim_Prefix_sfc, PublicConst.SQL_F_TRIM_PREFIX
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.Trim_Suffix_sfc, PublicConst.SQL_F_TRIM_SUFFIX
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.Connect_OR_sfc, PublicConst.SQL_F_CONNECT_OR
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.Connect_AND_sfc, PublicConst.SQL_F_CONNECT_AND
    dicSQLFieldCondition.Add Enum_SQL_F_Condition.Connect_Comma_sfc, PublicConst.SQL_F_CONNECT_COMMA
'Public Enum ACCDB_Data_Type
'    [Text] = 1
'    [INTEGER] = 2
'    [BIT] = 3
'    [Boolean] = 3
'    [COUNTER] = 4
'    [AUTOINCREMENT] = 4
'    [Decimal] = 5
'    Single_typ = 6
'    Double_Typ = 7
'End Enum
    If dicAccdbDataType Is Nothing Then
        Set dicAccdbDataType = New Dictionary
    End If
    dicAccdbDataType.Add ACCDB_Data_Type.Text_typ, "TEXT"
    dicAccdbDataType.Add ACCDB_Data_Type.Char_ShortText_typ, "CHAR"
    dicAccdbDataType.Add ACCDB_Data_Type.Integer_typ, "INTEGER"
    dicAccdbDataType.Add ACCDB_Data_Type.BIT_typ, "BIT"
    dicAccdbDataType.Add ACCDB_Data_Type.COUNTER_typ, "COUNTER"
    dicAccdbDataType.Add ACCDB_Data_Type.Decimal_typ, "DECIMAL"
    dicAccdbDataType.Add ACCDB_Data_Type.Single_typ, "REAL"
    dicAccdbDataType.Add ACCDB_Data_Type.Double_Typ, "FLOAT"
    'CATテーブルのEnum初期設定
    Call CAT_Enum_Initialset
    'INVテーブルのEnum初期設定
    Call INV_Enum_Initialize
End Sub
'''CATテーブルのEnumの初期化をする
'''---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CAT_Enum_Initialset()
'CATのヘッダテーブル定義のEnum、実際の値はclsEnumで定義する
'Public Enum Enum_CAT_Header
'    T_Name = 0              'Header Tableそのものの名前
'    F_Header = 1            '各CATコードのヘッダ部分
'    F_DescriptionTable = 2  'Description定義のテーブル名
'    F_DetailTable = 3       'Detailのテーブル名
'    F_SpecialTable = 4      'Specialのテーブル名
'    F_InputDate = 5         '入力日時
'End Enum
    If dicCATHeaderField Is Nothing Then
        Set dicCATHeaderField = New Dictionary
    End If
    dicCATHeaderField.Add Enum_CAT_Header.T_Name_chd, CAT_CONST.T_CAT_HEADER
    dicCATHeaderField.Add Enum_CAT_Header.F_Header_chd, CAT_CONST.F_CAT_HEADER
    dicCATHeaderField.Add Enum_CAT_Header.F_DescriptionTable_chd, CAT_CONST.F_CAT_DESCRIPTIONTABLE
    dicCATHeaderField.Add Enum_CAT_Header.F_DetailTable_chd, CAT_CONST.F_CAT_DETAILTABLE
    dicCATHeaderField.Add Enum_CAT_Header.F_SpecialTable_chd, CAT_CONST.F_CAT_SPECIALTABLE
'    dicCATHeaderField.Add Enum_CAT_Header.F_InputDate_chd, PublicConst.INPUT_DATE
''DescriptionテーブルのEnum
'Public Enum Enum_CAT_Description
'    T_Name_0_Kishu = 0          'インデックス0（機種名に置換が必要）Descriptionテーブル名
'    F_Digit_Row = 1             'フラグ立てたLongの数
'    F_Descriptoin_ID = 2        'DescriptionのID、実態はマスターテーブル参照
'    F_InputDate = 3             '入力日時
'End Enum
    If dicCATDescriptionField Is Nothing Then
        Set dicCATDescriptionField = New Dictionary
    End If
    dicCATDescriptionField.Add Enum_CAT_Description.T_Name_0_Kishu_cdc, CAT_CONST.T_CAT_DESCRIPTION_0kishu
    dicCATDescriptionField.Add Enum_CAT_Description.F_Digit_Row_cdc, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATDescriptionField.Add Enum_CAT_Description.F_Descriptoin_ID_cdc, CAT_CONST.F_CAT_DESCRIPTION_ID
'    dicCATDescriptionField.Add Enum_CAT_Description.F_InputDate_cdc, PublicConst.INPUT_DATE
''DetailテーブルのEnum
'Public Enum Enum_CAT_Detail
'    T_Name_0_Kishu = 0          'Detaiテーブル名 0 （機種名置換）が必要
'    F_Digit_Row = 1             '桁数フラグのLong
'    F_Chr = 2                   '桁に入る文字
'    F_Detail_ID = 3             'DetailのID
'    F_InputDate = 4             '入力日時
'End Enum
    If dicCATDetailField Is Nothing Then
        Set dicCATDetailField = New Dictionary
    End If
    dicCATDetailField.Add Enum_CAT_Detail.T_Name_0_Kishu_cdt, CAT_CONST.T_CAT_DETAIL_0kishu
    dicCATDetailField.Add Enum_CAT_Detail.F_Digit_Row_cdt, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATDetailField.Add Enum_CAT_Detail.F_Chr_cdt, CAT_CONST.F_CAT_CHR
    dicCATDetailField.Add Enum_CAT_Detail.F_Detail_ID_cdt, CAT_CONST.F_CAT_DETAIL_ID
'    dicCATDetailField.Add Enum_CAT_Detail.F_InputDate_cdt, PublicConst.INPUT_DATE
''SpecialテーブルのEnum
'Public Enum Enum_CAT_Special
'    T_Name_0_Kishu = 0          'Specialのテーブル名、{0}を機種名置換必要
'    F_Condition = 1             '条件フィールド
'    F_Execute = 2               '改変内容フィールド
'    F_InputDate = 3             '入力日時
'End Enum
    If dicCATSpecialField Is Nothing Then
        Set dicCATSpecialField = New Dictionary
    End If
    dicCATSpecialField.Add Enum_CAT_Special.T_Name_0_Kishu_csp, CAT_CONST.T_CAT_SPECIAL_0kishu
    dicCATSpecialField.Add Enum_CAT_Special.F_Condition_csp, CAT_CONST.F_CAT_CONDITION
    dicCATSpecialField.Add Enum_CAT_Special.F_Execute_csp, CAT_CONST.F_CAT_EXECUTE
'    dicCATSpecialField.Add Enum_CAT_Special.F_InputDate_csp, PublicConst.INPUT_DATE
''DescriptionマスターテーブルのEnum
'Public Enum Enum_CAT_M_Description
'    T_Name = 0                  'Descriptionマスターのテーブル名
'    F_Description_ID = 1        'Descriptionテーブルと共用、こっちが親
'    F_Digit_Row = 2             'Descriptionテーブルと共用
'    F_Description_Text = 3      'Descriptionの本体
'    F_InputDate = 4             '入力日時
'End Enum
    If dicCATMasterDescriptionField Is Nothing Then
        Set dicCATMasterDescriptionField = New Dictionary
    End If
    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.T_Name_cmdc, CAT_CONST.T_CAT_Description_MASTER
    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.F_Description_ID_cmdc, CAT_CONST.F_CAT_DESCRIPTION_ID
    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.F_Digit_Row_cmdc, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.F_Description_Text_cmdc, CAT_CONST.F_CAT_DESCRIPTION_TEXT
'    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.F_InputDate_cmdc, PublicConst.INPUT_DATE
''DetailマスターテーブルのEnum
'Public Enum Enum_CAT_M_Detail
'    T_Name = 0                  'Detailマスターのテーブル名
'    F_Detail_ID = 1             'Detailテーブルと共用、こっちが親
'    F_Digit_Row = 2             'Detailテーブルと共用
'    F_Chr = 3                   'Detailテーブルと共用
'    F_Detail_Text = 4           'Detailの本体
'    F_InputDate = 5             '入力日時
'End Enum
    If dicCATMasterDetailField Is Nothing Then
        Set dicCATMasterDetailField = New Dictionary
    End If
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.T_Name_cmdt, CAT_CONST.T_CAT_DETAIL_MASTER
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_Detail_ID_cmdt, CAT_CONST.F_CAT_DETAIL_ID
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_Digit_Row_cmdt, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_Chr_cmdt, CAT_CONST.F_CAT_CHR
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_Detail_Text_cmdt, CAT_CONST.F_CAT_DETAIL_TEXT
'    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_InputDate_cmdt, PublicConst.INPUT_DATE
''桁数マスターテーブルのEnum
'Public Enum Enum_CAT_M_Digit
'    T_Name = 0                  '桁数マスターテーブルの名前
'    F_Digit_Offset = 1          'ヘッダの文字の最後を0とした文字位置のオフセット（1文字目は1から始まる）
'    F_Digit_Row = 2             '他のテーブルと共用、こちらが親
'End Enum
    If dicMasterDigitField Is Nothing Then
        Set dicMasterDigitField = New Dictionary
    End If
    dicMasterDigitField.Add Enum_CAT_M_Digit.T_Name_cmdg, CAT_CONST.T_CAT_DIGIT_MASTER
    dicMasterDigitField.Add Enum_CAT_M_Digit.F_Digit_Offset_cmdg, CAT_CONST.F_CAT_DIGIT_OFFSET
    dicMasterDigitField.Add Enum_CAT_M_Digit.F_Digit_Row_cmdg, CAT_CONST.F_CAT_DIGIT_ROW
''一時利用フィールドEnum
'Public Enum CAT_Tmp
'    F_Digit_Update_ctm = 0
'End Enum
    If dicCATTmpField Is Nothing Then
        Set dicCATTmpField = New Dictionary
    End If
    dicCATTmpField.Add CAT_Tmp.F_Digit_Update_ctm, CAT_CONST.F_DIGIT_UPDATE
End Sub
Private Sub INV_Enum_Initialize()
    On Error GoTo ErrorCatch
    '在庫情報DBのEnumの初期化を行う
'手配コードマスターのEnum定義
'Public Enum Enum_INV_M_Parts
'    Table_Name_IMPrt = 1
'    F_Tehai_ID_IMPrt = 2
'    F_Tehai_Code_IMPrt = 3
'    F_Manege_Section_IMPrt = 4
'    F_System_TanaNo_IMPrt = 5
'    F_Kishu_IMPrt = 6
'    F_Store_Code_IMPrt = 7
'    F_Deliver_LOT_IMPrt = 8
'    F_Fill_Lot_IMPrt = 9
'    F_Lead_Time_IMPrt = 10
'    F_Order_Amount_IMPrt = 11
'    F_Order_Remain_IMPrt = 12
'    F_Stock_Amount_IMPrt = 13
'    F_Tana_ID_IMPrt = 14
'    F_System_Name_IMPrt = 15
'    F_System_Spec_IMPrt = 16
'    F_Store_Unit_IMPrt = 17
'    F_System_Description_IMPrt = 18
'    F_Local_Description_IMPrt = 19
'    F_Manege_Section_Sub_IMPrt = 20
'    F_InputDate_IMPrt = 21
'    'BINカード、表示関係でローカル限定
'    F_Label_Name_1_IMPrt = 101
'    F_Label_Name_2_IMPrt = 102
'    F_Label_Remark_1_IMPrt = 103
'    F_Label_Remark_2_IMPrt = 104
'End Enum
    If dicINVMasterParts Is Nothing Then
        Set dicINVMasterParts = New Dictionary
    End If
    dicINVMasterParts.Add Enum_INV_M_Parts.Table_Name_IMPrt, INV_CONST.T_INV_M_Parts
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Tehai_ID_IMPrt, INV_CONST.F_INV_TEHAI_ID
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Tehai_Code_IMPrt, INV_CONST.F_INV_TEHAI_TEXT
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Manege_Section_IMPrt, INV_CONST.F_INV_MANEGE_SECTON
    dicINVMasterParts.Add Enum_INV_M_Parts.F_System_TanaNo_IMPrt, INV_CONST.F_INV_SYSTEM_LOCATION_NO
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Kishu_IMPrt, INV_CONST.F_INV_KISHU
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Store_Code_IMPrt, INV_CONST.F_INV_STORE_CODE
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Deliver_Lot_IMPrt, INV_CONST.F_INV_DELIVER_LOT
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Fill_Lot_IMPrt, INV_CONST.F_INV_FILL_LOT
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Lead_Time_IMPrt, INV_CONST.F_INV_LEAD_TIME
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Order_Amount_IMPrt, INV_CONST.F_INV_ORDER_AMOUNT
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Order_Remain_IMPrt, INV_CONST.F_INV_ORDER_REMAIN
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Stock_Amount_IMPrt, INV_CONST.F_INV_STOCK_AMOUNT
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Tana_ID_IMPrt, INV_CONST.F_INV_TANA_ID
    dicINVMasterParts.Add Enum_INV_M_Parts.F_System_Name_IMPrt, INV_CONST.F_INV_SYSTEM_NAME
    dicINVMasterParts.Add Enum_INV_M_Parts.F_System_Spec_IMPrt, INV_CONST.F_INV_SYSTEM_SPEC
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Store_Unit_IMPrt, INV_CONST.F_INV_STORE_UNIT
    dicINVMasterParts.Add Enum_INV_M_Parts.F_System_Description_IMPrt, INV_CONST.F_INV_SYSTEM_DESCRIPTION
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Local_Description_IMPrt, INV_CONST.F_INV_LOCAL_DESCRIPTION
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Manege_Section_Sub_IMPrt, INV_CONST.F_INV_MANEGE_SECTION_SUB
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Label_Name_1_IMPrt, INV_CONST.F_INV_LABEL_NAME_1
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Label_Name_2_IMPrt, INV_CONST.F_INV_LABEL_NAME_2
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Label_Remark_1_IMPrt, INV_CONST.F_INV_LABEL_REMARK_1
    dicINVMasterParts.Add Enum_INV_M_Parts.F_Label_Remark_2_IMPrt, INV_CONST.F_INV_LABEL_REMARK_2
'    dicINVMasterParts.Add Enum_INV_M_Parts.F_InputDate_IMPrt, PublicConst.INPUT_DATE
''在庫情報シートのフィールド名（漢字）
''大部分をINV_Master_PartsのEnumと共有する
''棚番のみは共有できないので、独自に数字を振る
'Public Enum Enum_Sh_Zaiko
'    F_Tehai_Code_ShZ = Enum_INV_M_Parts.F_Tehai_Code_IMPrt
'    F_Manege_Section_ShZ = Enum_INV_M_Parts.F_Manege_Section_IMPrt
'    F_System_TanaNO_ShZ = Enum_INV_M_Parts.F_System_TanaNo_IMPrt
'    F_kishu_ShZ = Enum_INV_M_Parts.F_Kishu_IMPrt
'    F_Store_Code_ShZ = Enum_INV_M_Parts.F_Store_Code_IMPrt
'    F_Deliver_Lot_ShZ = Enum_INV_M_Parts.F_Deliver_Lot_IMPrt
'    F_Fill_Lot_ShZ = Enum_INV_M_Parts.F_Fill_Lot_IMPrt
'    F_Lead_Time_ShZ = Enum_INV_M_Parts.F_Lead_Time_IMPrt
'    F_Order_Amount_ShZ = Enum_INV_M_Parts.F_Order_Amount_IMPrt
'    F_Order_Remail_ShZ = Enum_INV_M_Parts.F_Order_Remain_IMPrt
'    F_Stock_Amount_ShZ = Enum_INV_M_Parts.F_Stock_Amount_IMPrt
'    F_System_Name_ShZ = Enum_INV_M_Parts.F_System_Name_IMPrt
'    F_System_Spec_ShZ = Enum_INV_M_Parts.F_System_Spec_IMPrt
'    F_Store_Unit_ShZ = Enum_INV_M_Parts.F_Store_Unit_IMPrt
'    F_System_Description_ShZ = Enum_INV_M_Parts.F_System_Description_IMPrt
'    F_Manege_Section_Sub_ShZ = Enum_INV_M_Parts.F_Manege_Section_Sub_IMPrt
'    '棚番テキストのみこちらで独自に設定する
'    F_Tana_Text_ShZ = 101
'End Enum
    If dicSHZaikoParts Is Nothing Then
        '在庫情報シートのDictionaryを初期化
        Set dicSHZaikoParts = New Dictionary
    End If
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Deliver_Lot_ShZ, INV_CONST.F_SH_ZAIKO_DELIVER_LOT
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Fill_Lot_ShZ, INV_CONST.F_SH_ZAIKO_FILL_LOT
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_kishu_ShZ, INV_CONST.F_SH_ZAIKO_KISHU
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Lead_Time_ShZ, INV_CONST.F_SH_ZAIKO_LEAD_TIME
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Manege_Section_ShZ, INV_CONST.F_SH_ZAIKO_MANEGE_SECTON
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Manege_Section_Sub_ShZ, INV_CONST.F_SH_ZAIKO_MANEGE_SECTION_SUB
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Order_Amount_ShZ, INV_CONST.F_SH_ZAIKO_ORDER_AMOUNT
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Order_Remain_ShZ, INV_CONST.F_SH_ZAIKO_ORDER_REMAIN
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Stock_Amount_ShZ, INV_CONST.F_SH_ZAIKO_STOCK_AMOUNT
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Store_Code_ShZ, INV_CONST.F_SH_ZAIKO_STORE_CODE
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Store_Unit_ShZ, INV_CONST.F_SH_ZAIKO_STORE_UNIT
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_System_Description_ShZ, INV_CONST.F_SH_ZAIKO_SYSTEM_DESCRIPTION
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_System_Name_ShZ, INV_CONST.F_SH_ZAIKO_SYSTEM_NAME
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_System_Spec_ShZ, INV_CONST.F_SH_ZAIKO_SYSTEM_SPEC
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_System_TanaNO_ShZ, INV_CONST.F_SH_ZAIKO_SYSTEM_TANA_NO
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Tana_Text_ShZ, INV_CONST.F_SH_ZAIKO_TANA_TEXT
    dicSHZaikoParts.Add Enum_Sh_Zaiko.F_Tehai_Code_ShZ, INV_CONST.F_SH_ZAIKO_TEHAI_TEXT
''T_M_Tanaフィールド定義Enum
'Public Enum Enum_INV_M_Tana
'    F_INV_TANA_ID = 1
'    F_INV_Tana_Local_Text = 2
'    F_INV_Tana_System_Text = 3
'    F_INV_TIET_Delivary = 4
'    F_InputDate = 5
'End Enum
If dicINVMasterTana Is Nothing Then
    Set dicINVMasterTana = New Dictionary
End If
'dicINVMasterTana.Add Enum_INV_M_Tana.F_InputDate_IMT, PublicConst.INPUT_DATE
    dicINVMasterTana.Add Enum_INV_M_Tana.F_INV_TANA_ID_IMT, INV_CONST.F_INV_TANA_ID
    dicINVMasterTana.Add Enum_INV_M_Tana.F_INV_Tana_Local_Text_IMT, INV_CONST.F_INV_TANA_LOCAL_TEXT
    dicINVMasterTana.Add Enum_INV_M_Tana.F_INV_Tana_System_Text_IMT, INV_CONST.F_INV_TANA_SYSTEM_TEXT
    dicINVMasterTana.Add Enum_INV_M_Tana.F_INV_TIET_Delivary_IMT, INV_CONST.F_INV_TANA_TIET_DELIVARY
''SQLAliasEnum
'Public Enum Enum_SQL_INV_Alias
'    INVDB_Parts_Alias_sia = 1
'    INVDB_Tana_Alias_sia = 2
'    INVDB_Tmp_Alias_sia = 3
'    ZaikoSH_Alias_sia = 4
'    TanaCSV_Alias_sia = 5
'    SHCSV_Alias_sia = 6
'    DstTable = 7
'    OriginTable = 8
'End Enum
    If dicSQLINVAlias Is Nothing Then
        Set dicSQLINVAlias = New Dictionary
    End If
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.INVDB_Parts_Alias_sia, INV_CONST.SQL_ALIAS_T_INVDB_Parts
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.INVDB_Tana_Alias_sia, INV_CONST.SQL_ALIAS_T_INVDB_Tana
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.INVDB_Tmp_Alias_sia, INV_CONST.SQL_ALIAS_T_TEMP
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.ZaikoSH_Alias_sia, SQL_ALIAS_T_SH_ZAIKO
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.TanaCSV_Alias_sia, SQL_ALIAS_T_INV_CSV
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.SHCSV_Alias_sia, INV_CONST.SQL_ALIAS_SH_CSV
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.DstTable_sia, INV_CONST.SQL_ALIAS_T_DST
    dicSQLINVAlias.Add Enum_SQL_INV_Alias.OriginTable_sia, INV_CONST.SQL_ALIAS_T_ORIG
''在庫情報シートでUpdate掛ける際にTrim必要なフィールド名を定義
''_ntrm need trim
'Public Enum Enum_SH_Zaiko_Need_Trim
'    F_Manege_Section_ntrm = Enum_Sh_Zaiko.F_Manege_Section_ShZ
'    F_Tehai_Code_ntrm = Enum_Sh_Zaiko.F_Tehai_Code_ShZ
'    F_System_TanaNO_ntrm = Enum_Sh_Zaiko.F_System_TanaNO_ShZ
'    F_Kishu_ntrm = Enum_Sh_Zaiko.F_kishu_ShZ
'    F_Store_Code_ntrm = Enum_Sh_Zaiko.F_Store_Code_ShZ
'    F_Tana_Text_ntrm = Enum_Sh_Zaiko.F_Tana_Text_ShZ
'    F_System_Name_ntrm = Enum_Sh_Zaiko.F_System_Name_ShZ
'    F_System_Spec_ntrm = Enum_Sh_Zaiko.F_System_Spec_ShZ
'    F_Store_Unit_ntrm = Enum_Sh_Zaiko.F_Store_Unit_ShZ
'    F_Manege_Section_sub_ntrm = Enum_Sh_Zaiko.F_Manege_Section_Sub_ShZ
'    F_System_Description_ntrm = Enum_Sh_Zaiko.F_System_Description_ShZ
'End Enum
    'NeedTrimだけ少し形式が違い、DictionaryにはEnumの定数のみ格納する
    If dicSHZaikoNeedTrim Is Nothing Then
        Set dicSHZaikoNeedTrim = New Dictionary
    End If
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Manege_Section_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Tehai_Code_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_System_TanaNO_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Kishu_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Store_Code_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Tana_Text_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_System_Name_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_System_Spec_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Store_Unit_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_Manege_Section_sub_ntrm, Empty
    dicSHZaikoNeedTrim.Add Enum_SH_Zaiko_Need_Trim.F_System_Description_ntrm, Empty
''CSV Enum Inv CSv file
''3個以外は共通なので共通(?)
'Public Enum Enum_CSV_Tana_Field
'    F_EndDay_ICS = 100
'    F_CSV_No_ICS = 101
'    F_Available_ICS = 102
'    'ロケーションだけメンドウなので独自に番号振る
'    F_Location_Text = 103
'    F_Status_ICS = 104
'    F_Bin_Amount_ICS = 105
'    '以下は共通フィールド
'    F_ManegeSection_ICS = Enum_INV_M_Parts.F_Manege_Section_IMPrt
'    F_ManegeSection_Sub_ICS = Enum_INV_M_Parts.F_Manege_Section_Sub_IMPrt
'    F_Tehai_Code_ICS = Enum_INV_M_Parts.F_Tehai_Code_IMPrt
'    F_System_Name_ICS = Enum_INV_M_Parts.F_System_Name_IMPrt
'    F_System_Spec_ICS = Enum_INV_M_Parts.F_System_Spec_IMPrt
'    F_System_Tana_NO_ICS = Enum_INV_M_Parts.F_System_TanaNo_IMPrt
'    F_Store_Code = Enum_INV_M_Parts.F_Store_Code_IMPrt
'    F_Stock_Amount = Enum_INV_M_Parts.F_Stock_Amount_IMPrt
'End Enum
    If dicCSVTana Is Nothing Then
        Set dicCSVTana = New Dictionary
    End If
    dicCSVTana.Add Enum_CSV_Tana_Field.F_EndDay_ICS, INV_CONST.F_INV_CSV_ENDDAY
    dicCSVTana.Add Enum_CSV_Tana_Field.F_CSV_No_ICS, INV_CONST.F_INV_CSV_NO
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Available_ICS, INV_CONST.F_INV_CSV_AVAILABLE_AMOUNT
    dicCSVTana.Add Enum_CSV_Tana_Field.F_ManegeSection_ICS, INV_CONST.F_INV_CSV_MANEGE_SECTION
    dicCSVTana.Add Enum_CSV_Tana_Field.F_ManegeSection_Sub_ICS, INV_CONST.F_INV_CSV_MANEGE_SECTION_SUB
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Tehai_Code_ICS, INV_CONST.F_INV_CSV_TEHAI_TEXT
    dicCSVTana.Add Enum_CSV_Tana_Field.F_System_Name_ICS, INV_CONST.F_INV_CSV_SYSTEM_NAME
    dicCSVTana.Add Enum_CSV_Tana_Field.F_System_Spec_ICS, INV_CONST.F_INV_CSV_SYSTEM_SPEC
    dicCSVTana.Add Enum_CSV_Tana_Field.F_System_Tana_NO_ICS, INV_CONST.F_INV_CSV_SYSTEM_TANA_NO
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Location_Text_ICS, INV_CONST.F_INV_CSV_LOCATION_TEXT
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Store_Code_ICS, INV_CONST.F_INV_CSV_STORE_CODE
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Stock_Amount_ICS, INV_CONST.F_INV_CSV_STOCK_AMOUNT
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Status_ICS, INV_CONST.F_INV_CSV_STATUS
    dicCSVTana.Add Enum_CSV_Tana_Field.F_Bin_Amount_ICS, INV_CONST.F_INV_CSV_BIN_AMOUNT
'dicCommonTanaCSV
    '特殊パターン、Tana_CSVファイルとDBのテーブルで共通の項目のEnumを格納したDictionary
    If dicCommonTanaCSV Is Nothing Then
        Set dicCommonTanaCSV = New Dictionary
    End If
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_EndDay_ICS, Empty        'AppendFieldで追加するので、EndDayも共通にしていい
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_CSV_No_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_Location_Text_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_ManegeSection_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_ManegeSection_Sub_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_Stock_Amount_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_Store_Code_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_System_Name_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_System_Spec_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_System_Tana_NO_ICS, Empty
    dicCommonTanaCSV.Add Enum_CSV_Tana_Field.F_Tehai_Code_ICS, Empty
'dicCSVTanaNeedTrim
    '特殊パターン、CSVにおいてtrimが必要なフィールドのEnumを格納する
    If dicCSVTanaNeedTrim Is Nothing Then
        Set dicCSVTanaNeedTrim = New Dictionary
    End If
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_EndDay_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_Location_Text_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_Manege_Section_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_Manege_Section_Sub_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_Store_Code_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_System_Name_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_System_Spec_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_System_Tana_No_ctrm, Empty
    dicCSVTanaNeedTrim.Add Enum_INV_CSV_Need_Trim.F_Tehai_Code_ctrm, Empty
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "INV_Enum_Initialize code: " & Err.Number & " Description: " & Err.Description
CloseAndExit:
    Exit Sub
End Sub
'DBExetension
Property Get DBFileExetension(DBext As DB_file_exetension) As String
    DBFileExetension = dicDBFileExetension(DBext)
End Property
'DBExetensionのDictionary
Property Get DBFileExetension_Dic() As Dictionary
    If dicDBFileExetension Is Nothing Then
        'メンバ変数がNothingだったらNothingを返す
        Set DBFileExetension_Dic = Nothing
    Else
        'メンバ変数のDictionaryよりKeyが拡張子（String)、ValueがEnumの定数（KeyとValueが反転する）を作成し、返す
        Dim dicLocalResult As Dictionary
        Set dicLocalResult = New Dictionary
        Dim varKey As Variant
        For Each varKey In dicDBFileExetension
            If Not dicLocalResult.Exists(dicDBFileExetension(varKey)) Then
                '拡張子がローカルdicのキーに存在しない場合のみ追加
                dicLocalResult.Add dicDBFileExetension(varKey), varKey
            End If
        Next varKey
        '作成したDictionaryを戻り値とする
        Set DBFileExetension_Dic = dicLocalResult
    End If
    Exit Property
End Property
'sqlFieldCondition
Property Get SQLFieldCondition(SQLCondition As Enum_SQL_F_Condition) As String
    SQLFieldCondition = dicSQLFieldCondition(SQLCondition)
    Exit Property
End Property
'AccdbDataType
Property Get AccdbDataType(accdbType As ACCDB_Data_Type) As String
    AccdbDataType = dicAccdbDataType(accdbType)
    Exit Property
End Property
'CATHeaderField
Property Get CATHeaderField(HeaderField As Enum_CAT_Header) As String
    CATHeaderField = dicCATHeaderField(HeaderField)
    Exit Property
End Property
'CATDescriptionField
Property Get CATDescriptionField(DescriptionField As Enum_CAT_Description) As String
    CATDescriptionField = dicCATDescriptionField(DescriptionField)
    Exit Property
End Property
'CATDetailField
Property Get CATDetailField(DetailField As Enum_CAT_Detail) As String
    CATDetailField = dicCATDetailField(DetailField)
    Exit Property
End Property
'CATSpecialField
Property Get CATSpecialField(SpecialField As Enum_CAT_Special) As String
    CATSpecialField = dicCATSpecialField(SpecialField)
    Exit Property
End Property
'MasterDescription
Property Get CATMasterDescriptionField(MasterDescriptionField As Enum_CAT_M_Description) As String
    CATMasterDescriptionField = dicCATMasterDescriptionField(MasterDescriptionField)
    Exit Property
End Property
'MasterDetail
Property Get CATMasterDetailField(MasterDetailField As Enum_CAT_M_Detail) As String
    CATMasterDetailField = dicCATMasterDetailField(MasterDetailField)
    Exit Property
End Property
'Digit
Property Get CATDigitField(DigitField As Enum_CAT_M_Digit) As String
    CATDigitField = dicMasterDigitField(DigitField)
    Exit Property
End Property
'CATTmp
Property Get CATTempField(TmpField As CAT_Tmp) As String
    CATTempField = dicCATTmpField(TmpField)
    Exit Property
End Property
'-----------------------------------------------------------------------------------------------------------------------------
'Sh_Zaiko
Property Get shZaikoField(ZaikoField As Enum_Sh_Zaiko) As String
    shZaikoField = dicSHZaikoParts(ZaikoField)
    Exit Property
End Property
'DicSHZaikoKeyParts
'SH_ZaikoのKeyとValueが反転したものを返す(KeyがFieldName)
Property Get Dic_ReverseSHZaiko() As Dictionary
    If dicSHZaikoKeyPartsValueEnum Is Nothing Then
        Set dicSHZaikoKeyPartsValueEnum = New Dictionary
        '反転DicがNothingだった場合、作成してやる
        Dim varSHKey As Variant
        For Each varSHKey In dicSHZaikoParts
            If Not dicSHZaikoKeyPartsValueEnum.Exists(dicSHZaikoParts(varSHKey)) Then
                '反転Dicに重複するkeyが無い場合のみ追加する
                'key value 反転
                dicSHZaikoKeyPartsValueEnum.Add dicSHZaikoParts(varSHKey), varSHKey
            End If
        Next varSHKey
    End If
    '反転Dicを返してやる
    Set Dic_ReverseSHZaiko = dicSHZaikoKeyPartsValueEnum
    Exit Property
End Property
'ZaikoSHNeedTrimDic
'在庫情報シートでTrimが必要なフィールドのEnumが入ったDictionaryを返す
Property Get Dic_NeedTrim_ZaikoSH() As Dictionary
    Set Dic_NeedTrim_ZaikoSH = dicSHZaikoNeedTrim
    Exit Property
End Property
'-----------------------------------------------------------------------------------------------------------------------------
'INV_DB
'INVMasterParts
Property Get INVMasterParts(MasterPartsField As Enum_INV_M_Parts) As String
    INVMasterParts = dicINVMasterParts(MasterPartsField)
    Exit Property
End Property
'INVMasterTana
Property Get INVMasterTana(MasterTana As Enum_INV_M_Tana) As String
    INVMasterTana = dicINVMasterTana(MasterTana)
    Exit Property
End Property
'T_INV_CSV
Property Get CSVTanafield(argCSVTanafield As Enum_CSV_Tana_Field) As String
    CSVTanafield = dicCSVTana(argCSVTanafield)
    Exit Property
End Property
'CSV_でkeyとvalueが反転したDictionaryを返す
Property Get Dic_Reverse_INV_CSV() As Dictionary
    If dic_ReverseCSVTana Is Nothing Then
        '中身が入ってなかったら新たにデータを作成する
        Set dic_ReverseCSVTana = New Dictionary
        Dim varKeyNormal As Variant
        For Each varKeyNormal In dicCSVTana
            'キーの重複確認
            If Not dic_ReverseCSVTana.Exists(dicCSVTana(varKeyNormal)) Then
                dic_ReverseCSVTana.Add dicCSVTana(varKeyNormal), varKeyNormal
            End If
        Next varKeyNormal
        '結果を返す
        Set Dic_Reverse_INV_CSV = dic_ReverseCSVTana
    Else
        '初期化済みなら中身が入ってるものとしてそのまま返す
        Set Dic_Reverse_INV_CSV = dic_ReverseCSVTana
        Exit Property
    End If
End Property
'CSVでNeedTrimフィールドのEnumを格納したDictionaryを返す
Property Get Dic_NeedTrim_CSV() As Dictionary
    Set Dic_NeedTrim_CSV = dicCSVTanaNeedTrim
    Exit Property
End Property
'SqlINV_Alias
Property Get SQL_INV_Alias(SqlAlias As Enum_SQL_INV_Alias) As String
    SQL_INV_Alias = dicSQLINVAlias(SqlAlias)
    Exit Property
End Property
'-----------------------------------------------------------------------------------------------------------------------------
'プロシージャー
'''Trimが必要なフィールドかどうかを返す
'''Return   bool                Trim必要なフィールドの時はTrue、それ以外はFalse
'''args
'''strargZaikoSHFieldName       外部から使用する必要があるので、引数はフィールド名そのもので受ける
'''argReversDic                 EnumのKey、Valueが反転したDictionaryを指定する（テーブル毎に存在）
Public Function IsNeedTrim(strargZaikoSHFieldName As String, ByRef argReversDic As Dictionary, ByRef argNeedtrimDic As Dictionary) As Boolean
    '引数チェック
    If strargZaikoSHFieldName = "" Then
        IsNeedTrim = False
        Exit Function
    End If
    If argReversDic Is Nothing Or argNeedtrimDic Is Nothing Then
        DebugMsgWithTime "IsNeedTrim: Reverse Dictionary empty"
        IsNeedTrim = False
    End If
    'そもそも指定シートのフィールド名かどうかを確認
    If Not argReversDic.Exists(strargZaikoSHFieldName) Then
        '反転dicに引数のフィールド名が見つからなかった
        DebugMsgWithTime "IsNeedRedim not found field name in SH_Zaiko Dictionary. Field Name: " & strargZaikoSHFieldName
        IsNeedTrim = False
        Exit Function
    End If
    '反転DicよりEnum数値を取り出し、それがNeedTrim の Dicに存在すればTrimが必要
    If argNeedtrimDic.Exists(argReversDic(strargZaikoSHFieldName)) Then
        '引数がNeedTrimのKeyに存在したらTrimが必要
        IsNeedTrim = True
    Else
        IsNeedTrim = False
    End If
    Exit Function
End Function
'''Author Daisuke Oota 2021_01_06
'''在庫情報シートとT_INV_Master_Partsの共通フィールドをStringの配列にして返す
'''戻り値 StringArray(n,1),nは共通フィールド数、2次元目は、0がDBフィールド、1がシートのフィールド名
'''xシートdiにしか無い値（手配コードとか）は2次元目は初期値になるので、受け取り側で判断する
'''共通項目のみ配列に入れるよう変更 2022_01_25 Patacchi
'''ex)(0,0) F_INV_Tehai_Code (0,1) 手配コード
Public Function GetUpdateCommonFieldZaikSHandMParts() As String()
    If dicINVMasterParts Is Nothing Or dicSHZaikoParts Is Nothing Then
        'DB側のdicかシート側のdicがNothingだった
        '何もせずに抜ける
        DebugMsgWithTime "GetUpdateCommonFieldZaikSHandMParts: member dictionary is nothing"
        Exit Function
    End If
    'シートdicでループ処理
'    Dim intArrRowCounter As Integer
'    intArrRowCounter = 0
    '結果を一旦Dictionaryに入れて、後で配列に展開する（片方しかデータなかった時のRedimが最終次元しか効かないため）
    'KeyにMarterPartsを、ValueにSHZaikoを入れる
    Dim varSHKey As Variant
    Dim dicResult As Dictionary
    Set dicResult = New Dictionary
    For Each varSHKey In dicSHZaikoParts
        If dicINVMasterParts.Exists(varSHKey) Then
            'シートdicと共通のKeyがDBdicにも存在する時
            'シートdicとDBdicそれぞれの値を結果dictionaryに入れてやる
            'DBに格納する時にTRIMが必要なものがあるので、切り分けを行う
            'ｘここでTrimはNG、後でテーブルエイリアスをプレフィックスでつけた時にダメになっちゃう
            Dim strPartsField As String
            Dim strSHField As String
            'DBPartsMaster
            strPartsField = dicINVMasterParts(varSHKey)
            'SHField
            strSHField = dicSHZaikoParts(varSHKey)
            dicResult.Add CStr(strPartsField), CStr(strSHField)
        Else
            'シートdicにしかkeyが存在しない場合
            '2次元目の1にシートdicのValueのみを入れてやる、2次元目の0は初期値
            'シートフィールドのみ
            '共通項目のみ返すよう変更
'            arrstrZaikoField(intArrRowCounter, 1) = CStr(dicSHZaikoParts(varshKey))
        End If
    Next varSHKey
    '結果を配列に格納しなおす
    Dim arrstrZaikoField() As String
    'String配列をシートdicの要素数で初期化
    ReDim arrstrZaikoField(dicResult.Count - 1, 1)
    Dim intArrayRowCounter
    intArrayRowCounter = 0
    For Each varSHKey In dicResult
        'DBフィールド
        arrstrZaikoField(intArrayRowCounter, 0) = CStr(varSHKey)
        'シートフィールド
        arrstrZaikoField(intArrayRowCounter, 1) = CStr(dicResult(varSHKey))
        '行カウンタをインクリメント
        intArrayRowCounter = intArrayRowCounter + 1
    Next varSHKey
    '結果を返す
    GetUpdateCommonFieldZaikSHandMParts = arrstrZaikoField
    Exit Function
End Function
'''INV_M_PartsとINV_M_Tanaの共通フィールドの配列を返す
'''JOINのONの後に続く文字列として使える
'''Return   String()    (Partsフィールド名,Tanaフィールド名)
Public Function GetKeyCommonFieldArrayINVPartsandINVTana() As String()
    'PartsとTanaは共通フィールドが一つしかないのが分かっているので、配列は静的に指定
    Dim arrstrCommonField(0, 1) As String
    '共通フィールドはF_INV_Tana_ID
    'Partテーブル
    arrstrCommonField(0, 0) = Me.INVMasterParts(F_Tana_ID_IMPrt)
    'Tanaテーブル
    arrstrCommonField(0, 1) = Me.INVMasterTana(F_INV_TANA_ID_IMT)
    GetKeyCommonFieldArrayINVPartsandINVTana = arrstrCommonField
    Exit Function
End Function
'''INV_M_Parts と Zaiko_SHの結合共通フィールドの配列を返す
'''JOIN の ON の後の文字列として使える
'''Return   String()    (Partsフィールド名,Zaiko_SHフィールド名)
Public Function GetKeyCommonFieldPatsAndZaikoSH() As String()
    '共通フィールドは複数あるが、JOINのKeyとしては手配コードとシステム棚番を使うので、配列は静的指定    2022_09_12 Daisuke oota
    '手配コードが共通で、システム棚番が違う物が出てきたので、ON条件は複合条件とする     2022_09_12 Daisuke oota
    '追加の条件は F_INV_Kishu フィールドとする
    Dim arrstrCommonZaikoSH(1, 1) As String
    '共通フィールドは手配コードと機種
    'Partsテーブル
    arrstrCommonZaikoSH(0, 0) = Me.INVMasterParts(F_Tehai_Code_IMPrt)
    arrstrCommonZaikoSH(1, 0) = Me.INVMasterParts(F_Kishu_IMPrt)
    'Zaiko_SHテーブル
    arrstrCommonZaikoSH(0, 1) = Me.shZaikoField(F_Tehai_Code_ShZ)
    arrstrCommonZaikoSH(1, 1) = Me.shZaikoField(F_kishu_ShZ)
    '結果を返す
    GetKeyCommonFieldPatsAndZaikoSH = arrstrCommonZaikoSH
    Exit Function
End Function
'''INV_M_Tana と Zaiko_SHの結合共通フィールドの配列を返す
'''JOIN の ON の後の文字列として使用
'''Return   String()    (Tanaフィールド,Zaiko_SHフィールド)
Public Function GetKeyCommonFieldTanaAndZaikoSH() As String()
    '結合キーフィールドはF_INV_Tana_System_Text
    Dim arrstrCommonTana(0, 1) As String
    'Tana
    arrstrCommonTana(0, 0) = Me.INVMasterTana(F_INV_Tana_System_Text_IMT)
    'Zako_SH
    arrstrCommonTana(0, 1) = Me.shZaikoField(F_Tana_Text_ShZ)
    '結果を返す
    GetKeyCommonFieldTanaAndZaikoSH = arrstrCommonTana
    Exit Function
End Function
'''T_INV_CSVと一時テーブルの対応テーブル配列を返す関数
'''Return   string()    2次元配列(DBフィールド名,Tempテーブルフィールド名)
'''でも棚CSVはテーブルとCSVファイルでフィールド名を共通にしたので、同じ文字が入るはず・・・・
'''プレフィックスはSQLSBクラスのメソッドで付加
'''共通フィールドのみ返す
Public Function GetCommonFieldArrayT_INV_CSV_andtemp() As String()
    'dicCommonTanaCSVに共通のフィールドのEnumが入ってるのでそれで・・・
    Dim strarrCommonCSVField() As String
    ReDim strarrCommonCSVField(dicCommonTanaCSV.Count - 1, 1)
    Dim varKeyCommon As Variant
    Dim longArrRowCounter As Long
    longArrRowCounter = 0
    For Each varKeyCommon In dicCommonTanaCSV
        'フィールド名をセットしていくが、この段階ではどちらも同じフィールド名が入る(DBをCSVのフィールド名で作成したため)
        strarrCommonCSVField(longArrRowCounter, 0) = Me.CSVTanafield(CLng(varKeyCommon))
        strarrCommonCSVField(longArrRowCounter, 1) = Me.CSVTanafield(CLng(varKeyCommon))
        longArrRowCounter = longArrRowCounter + 1
    Next varKeyCommon
    '得られた配列を結果として返す
    GetCommonFieldArrayT_INV_CSV_andtemp = strarrCommonCSVField
    Exit Function
End Function
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
Private dicSHZaikoParts As Dictionary
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
    dicCATHeaderField.Add Enum_CAT_Header.F_InputDate_chd, PublicConst.INPUT_DATE
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
    dicCATDescriptionField.Add Enum_CAT_Description.F_InputDate_cdc, PublicConst.INPUT_DATE
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
    dicCATDetailField.Add Enum_CAT_Detail.F_InputDate_cdt, PublicConst.INPUT_DATE
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
    dicCATSpecialField.Add Enum_CAT_Special.F_InputDate_csp, PublicConst.INPUT_DATE
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
    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.F_InputDate_cmdc, PublicConst.INPUT_DATE
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
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_InputDate_cmdt, PublicConst.INPUT_DATE
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
'End Enum
    If dicINVMasterParts Is Nothing Then
        Set dicINVMasterParts = New Dictionary
    End If
    dicINVMasterParts.Add Enum_INV_M_Parts.Table_Name_IMPrt, INV_CONST.T_INV_PARTS_MASTER
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
    dicINVMasterParts.Add Enum_INV_M_Parts.F_InputDate_IMPrt, PublicConst.INPUT_DATE
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
'-----------------------------------------------------------------------------------------------------------------------------
'INV_DB
'INVMasterParts
Property Get INVMasterParts(MasterPartsField As Enum_INV_M_Parts) As String
    INVMasterParts = dicINVMasterParts(MasterPartsField)
    Exit Property
End Property
'-----------------------------------------------------------------------------------------------------------------------------
'プロシージャー
'''Author Daisuke Oota 2021_01_06
'''在庫情報シートとT_INV_Master_Partsの共通フィールドをStringの配列にして返す
'''戻り値 StringArray(n,1),nは共通フィールド数、2次元目は、0がDBフィールド、1がシートのフィールド名
'''シートdiにしか無い値（手配コードとか）は2次元目は初期値になるので、受け取り側で判断する
'''ex)(0,0) F_INV_Tehai_Code (0,1) 手配コード
Public Function GetZaikoFieldNameArrayShandDB() As String()
    If dicINVMasterParts Is Nothing Or dicSHZaikoParts Is Nothing Then
        'DB側のdicかシート側のdicがNothingだった
        '何もせずに抜ける
        DebugMsgWithTime "GetZaikoFieldArray: member dictionary is nothing"
        Exit Function
    End If
    Dim arrZaikoField() As String
    'String配列をシートdicの要素数で初期化
    ReDim arrZaikoField(dicSHZaikoParts.Count - 1, 1)
    'シートdicでループ処理
    Dim intArrRowCounter As Integer
    intArrRowCounter = 0
    Dim varshKey As Variant
    For Each varshKey In dicSHZaikoParts
        If dicINVMasterParts.Exists(varshKey) Then
            'シートdicと共通のKeyがDBdicにも存在する時
            'シートdicとDBdicそれぞれの値をstring配列に入れてやる
            'DBフィールド
            arrZaikoField(intArrRowCounter, 0) = CStr(dicINVMasterParts(varshKey))
            'シートフィールド
            arrZaikoField(intArrRowCounter, 1) = CStr(dicSHZaikoParts(varshKey))
        Else
            'シートdicにしかkeyが存在しない場合
            '2次元目の1にシートdicのValueのみを入れてやる、2次元目の0は初期値
            'シートフィールドのみ
            arrZaikoField(intArrRowCounter, 1) = CStr(dicSHZaikoParts(varshKey))
        End If
        'StringArray行カウンターをインクリメント
        intArrRowCounter = intArrRowCounter + 1
    Next varshKey
    '結果を返す
    GetZaikoFieldNameArrayShandDB = arrZaikoField
    Exit Function
End Function
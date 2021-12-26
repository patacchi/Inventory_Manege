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
Private Sub Class_Initialize()
    Call Initial_Valueset
End Sub
Private Sub Initial_Valueset()
    'クラス初期化時にEnum毎のDictionaryに値を入れる
''DBファイルの列挙体
'Public Enum DB_file_exetension
'    accdb = 0
'End Enum
    If dicDBFileExetension Is Nothing Then
        Set dicDBFileExetension = New Dictionary
    End If
    dicDBFileExetension.Add DB_file_exetension.accdb_dext, "accdb"
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
    dicCATHeaderField.Add Enum_CAT_Header.F_InputDate_chd, CAT_CONST.INPUT_DATE
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
    dicCATDescriptionField.Add Enum_CAT_Description.F_InputDate_cdc, CAT_CONST.INPUT_DATE
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
    dicCATDetailField.Add Enum_CAT_Detail.F_InputDate_cdt, CAT_CONST.INPUT_DATE
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
    dicCATSpecialField.Add Enum_CAT_Special.F_InputDate_csp, CAT_CONST.INPUT_DATE
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
    dicCATMasterDescriptionField.Add Enum_CAT_M_Description.F_InputDate_cmdc, CAT_CONST.INPUT_DATE
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
    dicCATMasterDetailField.Add Enum_CAT_M_Detail.F_InputDate_cmdt, CAT_CONST.INPUT_DATE
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
'DBExetension
Property Get DBFileExetension(DBext As DB_file_exetension) As String
    DBFileExetension = dicDBFileExetension(DBext)
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
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
Private dicAccdbDataType As Dictionary
Private dicCATHeaderField As Dictionary
Private dicCATDescriptionField As Dictionary
Private dicCATDetailField As Dictionary
Private dicCATSpecialField As Dictionary
Private dicCATMasterDescriptionField As Dictionary
Private dicCATMasterDetailField As Dictionary
Private dicMasterDigitField As Dictionary
Private Sub Class_Initialize()
    Call Initial_Valueset
End Sub
Private Sub Initial_Valueset()
    'クラス初期化時にEnum毎のDictionaryに値を入れる
'Public Enum ACCDB_Data_Type
'    [Text] = 1
'    [INTEGER] = 2
'    [BIT] = 3
'    [Boolean] = 3
'    [COUNTER] = 4
'    [AUTOINCREMENT] = 4
'    [Decimal] = 5
'End Enum
    If dicAccdbDataType Is Nothing Then
        Set dicAccdbDataType = New Dictionary
    End If
    dicAccdbDataType.Add ACCDB_Data_Type.Text, "TEXT"
    dicAccdbDataType.Add ACCDB_Data_Type.Integer, "INTEGER"
    dicAccdbDataType.Add ACCDB_Data_Type.BIT, "BIT"
    dicAccdbDataType.Add ACCDB_Data_Type.COUNTER, "COUNTER"
    dicAccdbDataType.Add ACCDB_Data_Type.Decimal, "DECIMAL"
    'CATテーブルのEnum初期設定
    Call CAT_Enum_Initialset
End Sub
'''CATテーブルのEnumの初期化をする
'''---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CAT_Enum_Initialset()
'CATのヘッダテーブル定義のEnum、実際の値はclsEnumで定義する
'Public Enum CAT_Header_Field
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
    dicCATHeaderField.Add CAT_Header_Field.T_Name, CAT_CONST.T_CAT_HEADER
    dicCATHeaderField.Add CAT_Header_Field.F_Header, CAT_CONST.F_CAT_HEADER
    dicCATHeaderField.Add CAT_Header_Field.F_DescriptionTable, CAT_CONST.F_CAT_DESCRIPTIONTABLE
    dicCATHeaderField.Add CAT_Header_Field.F_DetailTable, CAT_CONST.F_CAT_DETAILTABLE
    dicCATHeaderField.Add CAT_Header_Field.F_SpecialTable, CAT_CONST.F_CAT_SPECIALTABLE
    dicCATHeaderField.Add CAT_Header_Field.F_InputDate, CAT_CONST.INPUT_DATE
''DescriptionテーブルのEnum
'Public Enum CAT_Description_Field
'    T_Name_0_Kishu = 0          'インデックス0（機種名に置換が必要）Descriptionテーブル名
'    F_Digit_Row = 1             'フラグ立てたLongの数
'    F_Descriptoin_ID = 2        'DescriptionのID、実態はマスターテーブル参照
'    F_InputDate = 3             '入力日時
'End Enum
    If dicCATDescriptionField Is Nothing Then
        Set dicCATDescriptionField = New Dictionary
    End If
    dicCATDescriptionField.Add CAT_Description_Field.T_Name_0_Kishu, CAT_CONST.T_CAT_DESCRIPTION_0kishu
    dicCATDescriptionField.Add CAT_Description_Field.F_Digit_Row, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATDescriptionField.Add CAT_Description_Field.F_Descriptoin_ID, CAT_CONST.F_CAT_DESCRIPTION_ID
    dicCATDescriptionField.Add CAT_Description_Field.F_InputDate, CAT_CONST.INPUT_DATE
''DetailテーブルのEnum
'Public Enum CAT_Detail_Field
'    T_Name_0_Kishu = 0          'Detaiテーブル名 0 （機種名置換）が必要
'    F_Digit_Row = 1             '桁数フラグのLong
'    F_Chr = 2                   '桁に入る文字
'    F_Detail_ID = 3             'DetailのID
'    F_InputDate = 4             '入力日時
'End Enum
    If dicCATDetailField Is Nothing Then
        Set dicCATDetailField = New Dictionary
    End If
    dicCATDetailField.Add CAT_Detail_Field.T_Name_0_Kishu, CAT_CONST.T_CAT_DETAIL_0kishu
    dicCATDetailField.Add CAT_Detail_Field.F_Digit_Row, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATDetailField.Add CAT_Detail_Field.F_Chr, CAT_CONST.F_CAT_CHR
    dicCATDetailField.Add CAT_Detail_Field.F_Detail_ID, CAT_CONST.F_CAT_DETAIL_ID
    dicCATDetailField.Add CAT_Detail_Field.F_InputDate, CAT_CONST.INPUT_DATE
''SpecialテーブルのEnum
'Public Enum CAT_Special_Field
'    T_Name_0_Kishu = 0          'Specialのテーブル名、{0}を機種名置換必要
'    F_Condition = 1             '条件フィールド
'    F_Execute = 2               '改変内容フィールド
'    F_InputDate = 3             '入力日時
'End Enum
    If dicCATSpecialField Is Nothing Then
        Set dicCATSpecialField = New Dictionary
    End If
    dicCATSpecialField.Add CAT_Special_Field.T_Name_0_Kishu, CAT_CONST.T_CAT_SPECIAL_0kishu
    dicCATSpecialField.Add CAT_Special_Field.F_Condition, CAT_CONST.F_CAT_CONDITION
    dicCATSpecialField.Add CAT_Special_Field.F_Execute, CAT_CONST.F_CAT_EXECUTE
    dicCATSpecialField.Add CAT_Special_Field.F_InputDate, CAT_CONST.INPUT_DATE
''DescriptionマスターテーブルのEnum
'Public Enum CAT_M_Description_Field
'    T_Name = 0                  'Descriptionマスターのテーブル名
'    F_Description_ID = 1        'Descriptionテーブルと共用、こっちが親
'    F_Digit_Row = 2             'Descriptionテーブルと共用
'    F_Description_Text = 3      'Descriptionの本体
'    F_InputDate = 4             '入力日時
'End Enum
    If dicCATMasterDescriptionField Is Nothing Then
        Set dicCATMasterDescriptionField = New Dictionary
    End If
    dicCATMasterDescriptionField.Add CAT_M_Description_Field.T_Name, CAT_CONST.T_CAT_Description_MASTER
    dicCATMasterDescriptionField.Add CAT_M_Description_Field.F_Description_ID, CAT_CONST.F_CAT_DESCRIPTION_ID
    dicCATMasterDescriptionField.Add CAT_M_Description_Field.F_Digit_Row, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATMasterDescriptionField.Add CAT_M_Description_Field.F_Description_Text, CAT_CONST.F_CAT_DESCRIPTION_TEXT
    dicCATMasterDescriptionField.Add CAT_M_Description_Field.F_InputDate, CAT_CONST.INPUT_DATE
''DetailマスターテーブルのEnum
'Public Enum CAT_M_Detail_Field
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
    dicCATMasterDetailField.Add CAT_M_Detail_Field.T_Name, CAT_CONST.T_CAT_DETAIL_MASTER
    dicCATMasterDetailField.Add CAT_M_Detail_Field.F_Detail_ID, CAT_CONST.F_CAT_DETAIL_ID
    dicCATMasterDetailField.Add CAT_M_Detail_Field.F_Digit_Row, CAT_CONST.F_CAT_DIGIT_ROW
    dicCATMasterDetailField.Add CAT_M_Detail_Field.F_Chr, CAT_CONST.F_CAT_CHR
    dicCATMasterDetailField.Add CAT_M_Detail_Field.F_Detail_Text, CAT_CONST.F_CAT_DETAIL_TEXT
    dicCATMasterDetailField.Add CAT_M_Detail_Field.F_InputDate, CAT_CONST.INPUT_DATE
''桁数マスターテーブルのEnum
'Public Enum CAT_M_Digit_Field
'    T_Name = 0                  '桁数マスターテーブルの名前
'    F_Digit_Offset = 1          'ヘッダの文字の最後を0とした文字位置のオフセット（1文字目は1から始まる）
'    F_Digit_Row = 2             '他のテーブルと共用、こちらが親
'End Enum
    If dicMasterDigitField Is Nothing Then
        Set dicMasterDigitField = New Dictionary
    End If
    dicMasterDigitField.Add CAT_M_Digit_Field.DigitT_Name, CAT_CONST.T_CAT_DIGIT_MASTER
    dicMasterDigitField.Add CAT_M_Digit_Field.F_Digit_Offset, CAT_CONST.F_CAT_DIGIT_OFFSET
    dicMasterDigitField.Add CAT_M_Digit_Field.F_DigitM_Row, CAT_CONST.F_CAT_DIGIT_ROW
End Sub
'AccdbDataType
Property Get AccdbDataType(accdbType As ACCDB_Data_Type) As String
    AccdbDataType = dicAccdbDataType(accdbType)
    Exit Property
End Property
'CATHeaderField
Property Get CATHeaderField(HeaderField As CAT_Header_Field) As String
    CATHeaderField = dicCATHeaderField(HeaderField)
    Exit Property
End Property
'CATDescriptionField
Property Get CATDescriptionField(DescriptionField As CAT_Description_Field) As String
    CATDescriptionField = dicCATDescriptionField(DescriptionField)
    Exit Property
End Property
'CATDetailField
Property Get CATDetailField(DetailField As CAT_Detail_Field) As String
    CATDetailField = dicCATDetailField(DetailField)
    Exit Property
End Property
'CATSpecialField
Property Get CATSpecialField(SpecialField As CAT_Special_Field) As String
    CATSpecialField = dicCATSpecialField(SpecialField)
    Exit Property
End Property
'MasterDescription
Property Get CATMasterDescriptionField(MasterDescriptionField As CAT_M_Description_Field) As String
    CATMasterDescriptionField = dicCATMasterDescriptionField(MasterDescriptionField)
    Exit Property
End Property
'MasterDetail
Property Get CATMasterDetailField(MasterDetailField As CAT_M_Detail_Field) As String
    CATMasterDetailField = dicCATMasterDetailField(MasterDetailField)
    Exit Property
End Property
'Digit
Property Get CATDigitField(DigitField As CAT_M_Digit_Field) As String
    CATDigitField = dicMasterDigitField(DigitField)
    Exit Property
End Property
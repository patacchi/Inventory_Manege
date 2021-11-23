VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCAT_SelectAlreadryKishu 
   Caption         =   "既存機種選択"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
   OleObjectBlob   =   "frmCAT_SelectAlreadryKishu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCAT_SelectAlreadryKishu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COLUMN_WIDTHSTR As String = "98.4pt"
Private Const LISTALREADY_COLUMN_COUNT As Long = 1
Private Sub UserForm_Activate()
    'リサイズ機能追加
    Call FormResize
End Sub
Private Sub UserForm_Initialize()
    'ヘッダテーブルより機種一覧を取得する
    Dim dbCAT As clsADOHandle
    Set dbCAT = New clsADOHandle
    dbCAT.DBFileName = CAT_CONST.CAT_DB_FILENAME
    dbCAT.SQL = CAT_CONST.SQL_KISHU_LIST
    Dim isCollect As Boolean
    isCollect = dbCAT.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        DebugMsgWithTime "Form Already Kishu Initialize Error"
        Exit Sub
    End If
    Dim varArr As Variant
    varArr = dbCAT.RS_Array
    Dim strWidths As String
    strWidths = GetColumnWidthString(varArr, boolMaxLengthFind:=False)
    With listBox_CAT_Already
'        .ColumnCount = UBound(varArr, 2) - LBound(varArr, 2) + 1
        'ヘッダ列のみ表示させる（他を非表示にしたい）
        .ColumnCount = LISTALREADY_COLUMN_COUNT
'        .ColumnWidths = strWidths
        .ColumnWidths = COLUMN_WIDTHSTR
        '.List = Join(varRetValue)
        .List = varArr
    End With
    Set dbCAT = Nothing
End Sub
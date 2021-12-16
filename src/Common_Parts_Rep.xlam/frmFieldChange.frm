VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFieldChange 
   Caption         =   "フィールドアップデート"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10095
   OleObjectBlob   =   "frmFieldChange.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFieldChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnDoUpdate_Click()
    'フィールド修正適用
    If lstBoxTable_Name.ListIndex = -1 Or lstBoxField_Name.ListIndex = -1 Then
        Exit Sub
    End If
    Dim adoFieldUpdate As clsADOHandle
    Set adoFieldUpdate = CreateclsADOHandleInstance
    '変更対象であるDigit_offsetフィールドが存在するかチェックする
    Dim isDigitOffset As Boolean
    isDigitOffset = adoFieldUpdate.IsFieldExists(lstBoxTable_Name.List(lstBoxTable_Name.ListIndex), F_CAT_DIGIT_OFFSET, _
                                                    txtBoxDB_FileName.Text, txtBoxDB_Directory.Text)
    'アップデートチェックフィールドが存在するかチェックする
    Dim isUpdateField As Boolean
    isUpdateField = adoFieldUpdate.IsFieldExists(lstBoxTable_Name.List(lstBoxTable_Name.ListIndex), F_DIGIT_UPDATE, _
                                                    txtBoxDB_FileName.Text, txtBoxDB_Directory.Text)
    'Digit_Rowフィールドが存在するかチェックする
    Dim isDigitRow As Boolean
    isDigitRow = adoFieldUpdate.IsFieldExists(lstBoxTable_Name.List(lstBoxTable_Name.ListIndex), F_CAT_DIGIT_ROW, _
                                                    txtBoxDB_FileName.Text, txtBoxDB_Directory.Text)
    'stop
    Stop
End Sub
Private Sub btnGetTableList_Click()
    Dim dbGetTable As clsADOHandle
    Set dbGetTable = New clsADOHandle
    dbGetTable.DBPath = txtBoxDB_Directory.Text
    dbGetTable.DBFileName = txtBoxDB_FileName.Text
    'テーブル一覧を取得
    Dim adoxcatChange As adox.Catalog
    Set adoxcatChange = New adox.Catalog
    adoxcatChange.ActiveConnection = dbGetTable.ConnectionString
    Dim adoxTable As adox.Table
    Dim strarrTableName() As String
    Dim longTableCount As Long
    longTableCount = 0
    For Each adoxTable In adoxcatChange.Tables
        If adoxTable.Type = "TABLE" Then
            ReDim Preserve strarrTableName(longTableCount)
            strarrTableName(longTableCount) = adoxTable.Name
            longTableCount = longTableCount + 1
        End If
    Next adoxTable
    lstBoxTable_Name.List = strarrTableName
    Set adoxcatChange = Nothing
    Set dbGetTable = Nothing
End Sub
Private Sub lstBOxField_Name_Change()
    'テーブル名とフィールド名どちらも選択されていたらUpdateボタンを有効にする
    If lstBoxField_Name.ListIndex >= 0 And lstBoxTable_Name.ListIndex >= 0 Then
        btnDoUpdate.Enabled = True
    Else
        btnDoUpdate.Enabled = False
    End If
End Sub
Private Sub lstBoxTable_Name_Change()
    '未選択状態なら抜ける
    If lstBoxTable_Name.ListIndex = -1 Then
        Exit Sub
    End If
    Dim dbFieldList As clsADOHandle
    Set dbFieldList = New clsADOHandle
    dbFieldList.DBPath = txtBoxDB_Directory.Text
    dbFieldList.DBFileName = txtBoxDB_FileName.Text
    Dim strSQL As String
    strSQL = "SELECT TOP 1 * FROM " & lstBoxTable_Name.List(lstBoxTable_Name.ListIndex)
    dbFieldList.SQL = strSQL
    Dim isCollect As Boolean
    isCollect = dbFieldList.Do_SQL_with_NO_Transaction()
    If Not isCollect Then
        Exit Sub
    End If
    If dbFieldList.RS.RecordCount <= 0 Then
        lstBoxField_Name.ListIndex = -1
        lstBoxField_Name.Clear
        Exit Sub
    End If
    Dim strarrFieldList() As String
    ReDim strarrFieldList(dbFieldList.RS.Fields.Count - 1)
    Dim longFieldCount As Long
    For longFieldCount = 0 To dbFieldList.RS.Fields.Count - 1
        strarrFieldList(longFieldCount) = dbFieldList.RS.Fields(longFieldCount).Name
    Next longFieldCount
    'フィールド名リストに配列を設定
    lstBoxField_Name.List = strarrFieldList
    'フィールド名リストを未選択状態にする
    lstBoxField_Name.ListIndex = -1
    'フィールド名を選択するまでUpdateボタンを無効に
    btnDoUpdate.Enabled = False
    Set dbFieldList = Nothing
End Sub
Private Sub UserForm_Initialize()
    '初期値を投入
    Dim dbChange As clsADOHandle
    Set dbChange = New clsADOHandle
    txtBoxDB_Directory.Text = dbChange.DBPath
    txtBoxDB_FileName.Text = dbChange.DBFileName
    txtBoxDate_Max.Text = GetLocalTimeWithMilliSec
    'テーブル一覧を取得
    Dim adoxcatChange As adox.Catalog
    Set adoxcatChange = New adox.Catalog
    adoxcatChange.ActiveConnection = dbChange.ConnectionString
    Dim adoxTable As adox.Table
    Dim strarrTableName() As String
    Dim longTableCount As Long
    longTableCount = 0
    For Each adoxTable In adoxcatChange.Tables
        If adoxTable.Type = "TABLE" Then
            ReDim Preserve strarrTableName(longTableCount)
            strarrTableName(longTableCount) = adoxTable.Name
            longTableCount = longTableCount + 1
        End If
    Next adoxTable
'    lstBoxTable_Name.List = strarrTableName
    Set adoxcatChange = Nothing
    Set dbChange = Nothing
End Sub
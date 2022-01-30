VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFieldChange 
   Caption         =   "フィールドアップデート"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12660
   OleObjectBlob   =   "frmFieldChange.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFieldChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SQL_DigitOffset_Update_Check_0_Tablename_1UpdateField As String = "SELECT * FROM {0} WHERE {1} = FALSE"
Private Sub btnDoUpdate_Click()
    'フィールド修正適用
    '全テーブルモードかどうかで処理を分岐する
    Select Case chkBoxAllTable.Value
    Case True
        '全テーブル一括モードの時
        'テーブル名リストボックスのテーブルをFor Eachで回す
        Dim varKey As Variant
        For Each varKey In lstBoxTable_Name.List
            Call DoUpdateTable(CStr(varKey))
        Next varKey
    Case False
        '通常モードの場合はテーブル名の選択は必須
        If lstBoxTable_Name.ListIndex = -1 Or lstBoxField_Name.ListIndex = -1 Then
            Exit Sub
        End If
        Call DoUpdateTable(lstBoxTable_Name.List(lstBoxTable_Name.ListIndex))
    End Select
End Sub
'''Author Patacchi 2021_12_26
'''テーブル名を引数として、テーブルの修正作業を行う
Private Sub DoUpdateTable(strargTableName As String)
    Dim adoFieldUpdate As clsADOHandle
    Set adoFieldUpdate = CreateclsADOHandleInstance
    'DBファイル名とディレクトリ名テキストボックスの内容をクラスインスタンスのプロパティにセットしてやる
    adoFieldUpdate.DBPath = txtBoxDB_Directory.Text
    adoFieldUpdate.DBFileName = txtBoxDB_FileName.Text
    'Enumクラスのインスタンスを利用してConstの数値を引っ張る
    Dim clsEnumValue As clsEnum
    Set clsEnumValue = CreateclsEnum
    'StringBuilderクラスのインスタンス
    Dim strBC As clsSQLStringBuilder
    Set strBC = CreateclsSQLStringBuilder
    'フィールド内容修正作業
    'DigitOffset
    '変更対象であるDigit_offsetフィールドが存在するかチェックする
    Dim isDigitOffset As Boolean
    isDigitOffset = adoFieldUpdate.IsFieldExists(strargTableName, clsEnumValue.CATDigitField(F_Digit_Offset_cmdg))
    If Not isDigitOffset Or strargTableName = clsEnumValue.CATDigitField(T_Name_cmdg) Then
        'そもそもDigitOffsetフィールドが無い場合
        'もしくはテーブル名がDigitMasterテーブルであった場合
        'アップデート対象のフィールドがないためDigitOffset関連の修正はしない
        DebugMsgWithTime "btnDoUpdate: DigitOffset field not found."
    Else
        'DigitOffset修正対象なので、修正続行
        'アップデートチェックフィールドが存在するかチェックする
        Dim isUpdateField As Boolean
        isUpdateField = adoFieldUpdate.IsFieldExists(strargTableName, clsEnumValue.CATTempField(F_Digit_Update_ctm))
        If Not isUpdateField Then
            'updateフィールドがなければ作成する
            Call adoFieldUpdate.AppendField(strargTableName, clsEnumValue.CATTempField(F_Digit_Update_ctm), Boolean_typ)
        End If
        'Digit_Rowフィールドが存在するかチェックする
        Dim isDigitRow As Boolean
        isDigitRow = adoFieldUpdate.IsFieldExists(strargTableName, clsEnumValue.CATDigitField(F_Digit_Row_cmdg))
        If Not isDigitRow Then
            'DigitRowフィールドがなければ作成する
            Call adoFieldUpdate.AppendField(strargTableName, clsEnumValue.CATDigitField(F_Digit_Row_cmdg), Integer_typ)
        End If
        'DigitOffsetフィールドデータ型をTextに
        Dim isCollect As Boolean
        isCollect = adoFieldUpdate.ChangeDataType(strargTableName, clsEnumValue.CATDigitField(F_Digit_Offset_cmdg), Text_typ, "(31)")
        If Not isCollect Then
            'データ型変更失敗
            DebugMsgWithTime "DoUpdateTable: fail change DataType"
            Exit Sub
        End If
        'DigitOffsetフィールド修正
'        CAT_CONST.SQL_FIX_DIGITOFFSET_0_TableName_1_DigitOffset_2_DigitRow_3_DigitUpdate
        'DigitOffset修正用パラメータDictionary設定
        Dim dicFixDigitOffset As Dictionary
        Set dicFixDigitOffset = New Dictionary
        '0_TableName_1_DigitOffset_2_DigitRow_3_DigitUpdate
        dicFixDigitOffset.Add 0, strargTableName
        dicFixDigitOffset.Add 1, clsEnumValue.CATDigitField(F_Digit_Offset_cmdg)
        dicFixDigitOffset.Add 2, clsEnumValue.CATDigitField(F_Digit_Row_cmdg)
        dicFixDigitOffset.Add 3, clsEnumValue.CATTempField(F_Digit_Update_ctm)
        'SQL作成
        Dim strSQL_FixDigitOffset As String
        strSQL_FixDigitOffset = strBC.ReplaceParm(CAT_CONST.SQL_FIX_DIGITOFFSET_0_TableName_1_DigitOffset_2_DigitRow_3_DigitUpdate, dicFixDigitOffset)
        'ConnectModeのWriteフラグの状態を調べる
        If Not adoFieldUpdate.ConnectMode And adModeWrite Then
            adoFieldUpdate.ConnectMode = adoFieldUpdate.ConnectMode Or adModeWrite
        End If
        'SQL実行
        isCollect = adoFieldUpdate.Do_SQL_with_NO_Transaction(strSQL_FixDigitOffset)
        If Not isCollect Then
            'DigitRow設定時に失敗したっぽい
            DebugMsgWithTime "DoUpdateTable: fail Fix DigitRow"
            Exit Sub
        End If
        '修正完了確認
        'UpdatedでFalseが残ってないかどうか
        '置換パラメータ用Dictionary初期化
        dicFixDigitOffset.RemoveAll
        '0_Tablename 1_UpdateFieldName
        '置換用Dictionary作成
        dicFixDigitOffset.Add 0, strargTableName
        dicFixDigitOffset.Add 1, clsEnumValue.CATTempField(F_Digit_Update_ctm)
        strSQL_FixDigitOffset = strBC.ReplaceParm(SQL_DigitOffset_Update_Check_0_Tablename_1UpdateField, dicFixDigitOffset)
        Call adoFieldUpdate.Do_SQL_with_NO_Transaction(strSQL_FixDigitOffset)
        If adoFieldUpdate.RecordCount >= 1 Then
            'UpdateでFalseがまだ残ってる
            'メッセージボックス出して処理を中断
            DebugMsgWithTime "DoUpdateTable Table: " & strargTableName & " Update update NOT complete.check master table"
            MsgBox strargTableName & " テーブルでアップデート完了していないフィールドが残っているようです。マスターファイルを確認して下さい。"
        Else
            '全てアップデート完了した
            'DigitOffsetフィールド消去
            Call adoFieldUpdate.DeleteField(strargTableName, clsEnumValue.CATDigitField(F_Digit_Offset_cmdg))
        End If
    End If
    'InputDateフィールドがあるかチェックする
    Dim isInputDate As Boolean
    isInputDate = adoFieldUpdate.IsFieldExists(strargTableName, clsEnumValue.CATMasterDetailField(F_InputDate_cmdt))
    If isInputDate Then
        '修正対象のInputDateフィールドがある場合のみ修正作業を続行する
        '置換パラメータ用のDictionary作成
        Dim dicFixInputDate As Dictionary
        Set dicFixInputDate = New Dictionary
        dicFixInputDate.Add 0, strargTableName
        Dim strSQL_FixInputDate As String
        strSQL_FixInputDate = strBC.ReplaceParm(CAT_CONST.SQL_FIX_INPUTDATE_0_TableName, dicFixInputDate)
        '修正実行
        Call adoFieldUpdate.Do_SQL_with_NO_Transaction(strSQL_FixInputDate)
        'Writeフラグを下げる
        adoFieldUpdate.ConnectMode = adoFieldUpdate.ConnectMode And Not adModeWrite
    End If
    Set clsEnumValue = Nothing
    Set adoFieldUpdate = Nothing
    Set strBC = Nothing
End Sub
Private Sub btnGetTableList_Click()
    Dim dbGetTable As clsADOHandle
    Set dbGetTable = New clsADOHandle
    dbGetTable.DBPath = txtBoxDB_Directory.Text
    dbGetTable.DBFileName = txtBoxDB_FileName.Text
    'DBPathとDBFilenameテキストボックスが削除されている場合は標準設定を引っ張ってるはずなので、クラスのプロパティの内容をテキストボックスに反映してやる
    txtBoxDB_Directory.Text = dbGetTable.DBPath
    txtBoxDB_FileName.Text = dbGetTable.DBFileName
    'DBFileの存在有無確認
    Dim isDBFileExists As Boolean
    isDBFileExists = dbGetTable.IsDBFileExist(dbGetTable.DBFileName, dbGetTable.DBPath)
    If Not isDBFileExists Then
        'ファイルが存在しなかったら抜ける
        MsgBox "btnGetTableList_Click Path:  " & dbGetTable.DBPath & vbCrLf & " Filename: " & dbGetTable.DBFileName & " is not exists"
        Exit Sub
    End If
    'テーブル一覧を取得
    Dim adoxcatChange As ADOX.Catalog
    Set adoxcatChange = New ADOX.Catalog
    adoxcatChange.ActiveConnection = dbGetTable.ConnectionString
    Dim adoxTable As ADOX.Table
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
    If lstBoxTable_Name.ListCount >= 1 Then
        'テーブル名リストボックスにデータがあった場合、全テーブルボタンのVisibleをTrueにしてやる
        chkBoxAllTable.Visible = True
    Else
        'テーブルが見つからなかった場合は、全テーブルボタンのVisibleをFalseにしてやる
        chkBoxAllTable.Visible = False
    End If
    Set adoxcatChange = Nothing
    Set dbGetTable = Nothing
End Sub
Private Sub chkBoxAllTable_Change()
    '全テーブル適用チェックボタンの状態が変化した場合に実行
    Select Case chkBoxAllTable.Value
    Case True
        '全テーブルモードの場合
        'テーブル選択リストボックスとフィールド選択リストボックスのEnableをFalseにし、未選択状態にする
        lstBoxTable_Name.Enabled = False
        lstBoxTable_Name.ListIndex = -1
        lstBoxField_Name.Enabled = False
        lstBoxField_Name.ListIndex = -1
        'アップデートボタンを有効に
        btnDoUpdate.Enabled = True
    Case False
        '通常モード
        'リストボックスを有効にしてやる
        lstBoxTable_Name.Enabled = True
        lstBoxField_Name.Enabled = True
        'テーブル名リストボックスが未選択状態ならアップデートボタンを無効にしてやる
        If lstBoxTable_Name.ListIndex = -1 Then
            'テーブル名が未選択状態
            btnDoUpdate.Enabled = False
        End If
    End Select
End Sub
Private Sub IEbrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    '本来はWebページを表示するコントロールだが、ドラッグアンドドロップされるとURLにファイル名がそのまま入るため、これを利用する
    '複数ファイルの選択は不可
    '実際にNavigateを実行しないようにCancelにTrueにセットする
    Cancel = True
    Dim fsoFileNameGet As FileSystemObject
    Set fsoFileNameGet = New FileSystemObject
    If Not fsoFileNameGet.FileExists(URL) Then
        'URLのファイル名が存在しなかった→ファイル以外がドロップされた可能性があるので即抜ける
        DebugMsgWithTime "IEBrowser_BeforeNavigate: file?: " & URL & " not found"
        Exit Sub
    End If
    '拡張子よりDBのファイルかどうかを判定する
    Dim adoExtention As clsADOHandle
    Set adoExtention = CreateclsADOHandleInstance
    If Not adoExtention.IsDBExtention(CStr(URL)) Then
        'DBファイルの拡張子では無かった場合
        '何もしないで抜ける
        DebugMsgWithTime "GetFileName_Brouser: Target file is not DB file " & URL
        GoTo CloseAndExit
        Exit Sub
    End If
    'DBファイル名とディレクトリに設定してやる
    txtBoxDB_Directory.Text = fsoFileNameGet.GetParentFolderName(URL)
    txtBoxDB_FileName.Text = fsoFileNameGet.GetFileName(URL)
    GoTo CloseAndExit
CloseAndExit:
    Set adoExtention = Nothing
    Set fsoFileNameGet = Nothing
    Exit Sub
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
    'Adoxを利用して、フィールド一覧を取得する
    Dim adoxCatField As ADOX.Catalog
    Set adoxCatField = New ADOX.Catalog
    adoxCatField.ActiveConnection = dbFieldList.ConnectionString
    If adoxCatField.Tables(lstBoxTable_Name.List(lstBoxTable_Name.ListIndex)).Columns.Count >= 1 Then
        '列数のカウントが1以上の場合
        '既存のフィールド名リストボックスを消去
        lstBoxField_Name.Clear
        lstBoxField_Name.ListIndex = -1
        Dim adoxColumn As ADOX.Column
        For Each adoxColumn In adoxCatField.Tables(lstBoxTable_Name.List(lstBoxTable_Name.ListIndex)).Columns
            'リストボックスに列名を追加する
            lstBoxField_Name.AddItem adoxColumn.Name
        Next adoxColumn
        Set adoxCatField = Nothing
    Else
        '列数がなかった場合
        lstBoxField_Name.ListIndex = -1
        lstBoxField_Name.Clear
        lstBoxField_Name.AddItem ("レコード件数が0件以下でした。")
        'レコードなしの場合はフィールドリストボックスを無効にする
        btnDoUpdate.Enabled = False
        lstBoxField_Name.Enabled = False
        Exit Sub
    End If
    btnDoUpdate.Enabled = False
    Set dbFieldList = Nothing
End Sub
Private Sub txtBoxDB_Directory_Change()
    'ディレクトリ名が変化したとき
    'リストボックスを初期化する
    Call ClearTableandFieldList
End Sub
Private Sub txtBoxDB_FileName_Change()
    'ファイル名が変化したとき
    'リストボックスを初期化する
    Call ClearTableandFieldList
End Sub
'''Author Disuke Oota 2021_12_19
'''初期化用、テーブルリストとフィールドリストを消し去る
Private Sub ClearTableandFieldList()
    'テーブルリストの選択状態を解除し、リストを消去する
    lstBoxTable_Name.ListIndex = -1
    lstBoxTable_Name.Clear
    'フィールドリストの選択状態を解除し、リストを消去する
    lstBoxField_Name.ListIndex = -1
    lstBoxField_Name.Clear
    '全テーブルボタンのVisibleをFalseにしてやる
    chkBoxAllTable.Visible = False
End Sub
Private Sub UserForm_Initialize()
    '初期値を投入
    Dim dbChange As clsADOHandle
    Set dbChange = New clsADOHandle
    txtBoxDB_Directory.Text = dbChange.DBPath
    txtBoxDB_FileName.Text = dbChange.DBFileName
    txtBoxDate_Max.Text = GetLocalTimeWithMilliSec
    'テーブル一覧を取得
    Dim adoxcatChange As ADOX.Catalog
    Set adoxcatChange = New ADOX.Catalog
    adoxcatChange.ActiveConnection = dbChange.ConnectionString
    Dim adoxTable As ADOX.Table
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
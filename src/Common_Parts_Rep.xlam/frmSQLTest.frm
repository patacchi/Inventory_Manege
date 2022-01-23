VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQLTest 
   Caption         =   "SQLテスト"
   ClientHeight    =   8865.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13785
   OleObjectBlob   =   "frmSQLTest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSQLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Browser_GetFileName_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
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
        Exit Sub
    End If
'    Dim EnumValue As clsEnum
'    Set EnumValue = CreateclsEnum
'    If Not LCase(fsoFileNameGet.GetExtensionName(URL)) = LCase(EnumValue.DBFileExetension(accdb_dext)) Then
'        'ファイルの拡張子がDBFileExtentionと一致しない場合は処理を中断する
'        DebugMsgWithTime "IEBrowser_BeforeNavigate: Exetention is not accdb"
'        Exit Sub
'    End If
    'DBファイル名とディレクトリに設定してやる
    txtBoxDefaultDBDirectory.Text = fsoFileNameGet.GetParentFolderName(URL)
    txtBoxDefaultDBFile.Text = fsoFileNameGet.GetFileName(URL)
End Sub
Private Sub btnExportCSV_Click()
    'CSV出力
    Dim strFilePath As String
    strFilePath = Application.GetSaveAsFilename(InitialFileName:="\\PC24929-tdms\DBLearn\Test\CSV_Output\", filefilter:="CSVファイル(*.csv),*.csv")
    If strFilePath = "False" Then
        DebugMsgWithTime "btnExportCSVでキャンセルが押された"
        Exit Sub
    End If
    Call OutputArrayToCSV(Me.listBoxSQLResult.List, strFilePath)
    Exit Sub
End Sub
'''Author Daisuke oota 2021_10_29
''' 単体でテストしたいプロシージャを記述
'''
Private Sub btnSingleTest_Click()
#If DebugDB Then
    MsgBox "DebugDB = 1"
#End If
    'オートフィルタ設定・確認
    Dim InvDBTest As clsINVDB
    Set InvDBTest = CreateclsINVDB
    Dim fsoDBTest As FileSystemObject
    Set fsoDBTest = New FileSystemObject
    'クラスのプロパティにExcelファイルのフルパスを設定
    InvDBTest.BKZAikoInfoFullPath = fsoDBTest.BuildPath(txtBoxDefaultDBDirectory.Text, txtBoxDefaultDBFile.Text)
    'フィルタ処理し、結果の範囲の名前を受け取る
    Dim arrstrRangeName() As String
'    arrstrRangeName = InvDBTest.GetFilterRangeNameFromExcel
    Dim adoSingle As clsADOHandle
    Set adoSingle = CreateclsADOHandleInstance
    Dim clsEnumSingle As clsEnum
    Set clsEnumSingle = CreateclsEnum
    '手配コード最初4桁取得テスト
    Dim arrstrResult() As String
    arrstrResult = InvDBTest.Return4digitTehaiCodeFromCSV("", adoSingle, InvDBTest, clsEnumSingle)
    MsgBox "先頭4文字の種類は: " & CStr(UBound(arrstrResult) + 1)
'    'Select INTO テスト
'    Dim adoSingle As clsADOHandle
'    Set adoSingle = CreateclsADOHandleInstance
'    Dim clsEnumSingle As clsEnum
'    Set clsEnumSingle = CreateclsEnum
'    MsgBox "変更箇所は:" & CStr(InvDBTest.UpsertINVPartsMasterfromZaikoSH(arrstrRangeName(0, 0), InvDBTest, adoSingle, clsEnumSingle))
'    GoTo CloseAndExit
CloseAndExit:
    Set clsEnumSingle = Nothing
    Set InvDBTest = Nothing
    Set fsoDBTest = Nothing
    Exit Sub
'    Dim logBeki As Double
'''''32ビットまで順番にフラグを立てて、Longでどう表現されるか
''
''    Dim longFlag As Long
''    Dim intBitCount As Integer
''    Dim logBeki As Double
''    longFlag = 0
''    For intBitCount = 0 To 30
''        longFlag = 0 Or (2 ^ intBitCount)
''        logBeki = Log(longFlag) / Log(2#)
''        DebugMsgWithTime (vbCrLf & intBitCount & "bit" & vbCrLf & longFlag & vbCrLf & logBeki)
''    Next intBitCount
''ダウンロードパス取得
'    MsgBox GetDownloadPath
End Sub
'''Author Daisuke Oota 2021_10_18
'''パラメータバインドを使用するかどうか
'''Trueになったらパラメータ入力ボックスを表示させる、Falseになったら消す
Private Sub chkBoxUseParm_Change()
    Select Case chkBoxUseParm.Value
    Case True
        'パラメータバインドを使用する場合
        txtBoxParmText1.Visible = True
        txtBoxParmText2.Visible = True
        txtBoxParmText3.Visible = True
    Case False
        'パラメータバインドを使用しない場合
        txtBoxParmText1.Visible = False
        txtBoxParmText2.Visible = False
        txtBoxParmText3.Visible = False
    End Select
End Sub
'''Author Daisuke oota 2021_10_18
'''テキストボックスの値よりパラメータバインドに使用する置換リスト（Dictionary）を作成する
'''---------------------------------------------------------------------------------------------------
Private Function CreateParmDic() As Dictionary
    If txtBoxParmText1.Text = "" And txtBoxParmText2.Text = "" And txtBoxParmText3.Text = "" Then
        MsgBox "パラメータ入力テキストボックスが全て空です"
        Exit Function
    End If
    Dim localDic As Dictionary
    Set localDic = New Dictionary
    localDic.Add 0, txtBoxParmText1.Text
    localDic.Add 1, txtBoxParmText2.Text
    localDic.Add 2, txtBoxParmText3.Text
    Set CreateParmDic = localDic
    Exit Function
End Function
Private Sub UserForm_Activate()
    'リサイズ機能追加
    Call FormResize
End Sub
Private Sub UserForm_Initialize()
    'デフォルトDBディレクトリとDBファイル名を拾ってくる
    Dim dbDefault As clsADOHandle
    Set dbDefault = New clsADOHandle
    txtBoxDefaultDBDirectory.Text = dbDefault.DBPath
    txtBoxDefaultDBFile.Text = dbDefault.DBFileName
    '途中で簡単にディレクトリとファイルを切り替えれるようになったのでテキストボックスのEnableをTrueにセットしてやる
    txtBoxDefaultDBDirectory.Enabled = True
    txtBoxDefaultDBFile.Enabled = True
End Sub
Private Sub UserForm_Resize()
    'フォームリサイズ時に、中のリストボックスもサイズ変更してやる
    Dim intListHeight As Integer
    Dim intListWidth As Integer
    intListHeight = Me.InsideHeight - listBoxSQLResult.Top * 2
    intListWidth = Me.InsideWidth - (txtboxSQLText.Left * 2) - txtboxSQLText.Width - (listBoxSQLResult.Left - txtboxSQLText.Width - txtboxSQLText.Left)
    If (intListHeight > 0 And intListWidth > 0) Then
        listBoxSQLResult.Height = intListHeight
        listBoxSQLResult.Width = intListWidth
    End If
End Sub
Private Sub btnBulkDataInput_Click()
    Dim strSQL
    Randomize
'    frmBulkInsertTest.Show
    'ある範囲の乱数の発生のさせ方
    'Int((範囲上限値 - 範囲下限値 + 1) * Rnd + 範囲下限値)
End Sub
Private Sub btnSQLGo_Click()
    'エラーチェックとかほとんどなし
    'テキストボックスに入れたSQLを実行するフォームっぽいの
    If txtboxSQLText.Text = "" Then
        MsgBox "空白はちょっと・・・"
        Exit Sub
    End If
    Dim varRetValue As Variant
    Dim strWidths As String
    Dim isCollect As Boolean
    Dim dbTest As clsADOHandle
    Set dbTest = New clsADOHandle
    'DBディレクトリ・DBファイル名テキストボックス名で指定されたファイルがあるかチェックする
    Dim IsDBFileExist As Boolean
    IsDBFileExist = dbTest.IsDBFileExist(txtBoxDefaultDBFile.Text, txtBoxDefaultDBDirectory.Text)
    If Not IsDBFileExist Then
        MsgBox "DB directory: " & txtBoxDefaultDBDirectory.Text & " Filename: " & txtBoxDefaultDBFile.Text & " が見つかりませんでした。"
        Exit Sub
    End If
    'テキストボックスで指定したディレクトリ名とファイル名をクラスのプロパティにセットしてやる
    dbTest.DBPath = txtBoxDefaultDBDirectory.Text
    dbTest.DBFileName = txtBoxDefaultDBFile.Text
    If chkBoxUseParm.Value Then
        'パラメータバインド有りの場合
        Dim sqlBC As clsSQLStringBuilder
        Set sqlBC = New clsSQLStringBuilder
        Dim dicParm As Dictionary
        Set dicParm = CreateParmDic
        isCollect = dbTest.Do_SQL_with_NO_Transaction(sqlBC.ReplaceParm(txtboxSQLText.Text, dicParm))
        Set dicParm = Nothing
        Set sqlBC = Nothing
    Else
        isCollect = dbTest.Do_SQL_with_NO_Transaction(txtboxSQLText.Text)
    End If
    If isCollect Then
        If chkboxNoTitle.Value = True Then
            'タイトルなしを希望の場合はこちら
'            varRetValue = dbSQLite3.RS_Array(boolPlusTytle:=False)
            varRetValue = dbTest.RS_Array
            strWidths = GetColumnWidthString(varRetValue, 0)
        Else
            'デフォルトはタイトルあり
'            varRetValue = dbSQLite3.RS_Array(boolPlusTytle:=True)
            varRetValue = dbTest.RS_Array
            strWidths = GetColumnWidthString(varRetValue, 1)
        End If
    Else
        'エラーがあった場合の処理・・・なんだけど
        'エラーメッセージをそのまま表示すればいいのでは・・・
        If chkboxNoTitle.Value = True Then
            'タイトルなしを希望の場合はこちら
'            varRetValue = dbSQLite3.RS_Array(boolPlusTytle:=False)
            strWidths = GetColumnWidthString(varRetValue, 0)
        Else
            'デフォルトはタイトルあり
'            varRetValue = dbSQLite3.RS_Array(boolPlusTytle:=True)
            strWidths = GetColumnWidthString(varRetValue, 1)
        End If
    End If
    If VarType(varRetValue) = vbEmpty Then
        listBoxSQLResult.Clear
        listBoxSQLResult.AddItem "データなし"
        Exit Sub
    End If
    If chkBoxMaxLength.Value = True Then
        '最大文字数検索をしたいそうで
        strWidths = GetColumnWidthString(varRetValue, boolMaxLengthFind:=True)
    End If
    With listBoxSQLResult
        .ColumnCount = UBound(varRetValue, 2) - LBound(varRetValue, 2) + 1
        .ColumnWidths = strWidths
        '.List = Join(varRetValue)
        .List = varRetValue
        '.AddItem (varRetValue(1)(1))
    End With
End Sub
Private Sub listBoxSQLResult_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'リストダブルクリックしたらクリップボードにコピーしてみおよう
    Dim objDataObj As DataObject
    Dim intCounterColumn As Integer
    Dim strListText As String
    Set objDataObj = New DataObject
        objDataObj.SetText (listBoxSQLResult.List(listBoxSQLResult.ListIndex))
        objDataObj.PutInClipboard
        strListText = ""
        For intCounterColumn = 0 To listBoxSQLResult.ColumnCount - 1
            If IsNull(listBoxSQLResult.List(listBoxSQLResult.ListIndex, intCounterColumn)) Then
                'Nullの場合はNULLって入れてやろう
                strListText = strListText & " NULL"
            Else
                strListText = strListText & " " & CStr(listBoxSQLResult.List(listBoxSQLResult.ListIndex, intCounterColumn))
            End If
        Next intCounterColumn
        LTrim (strListText)
        MsgBox strListText
        DebugMsgWithTime strListText
End Sub
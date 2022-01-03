VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQLTest 
   Caption         =   "SQLテスト"
   ClientHeight    =   8865
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
    'オートフィルタ設定・確認
    'Excelファイルかどうか確認する
    Dim fsoFilter As FileSystemObject
    Set fsoFilter = New FileSystemObject
    Dim EnumValue As clsEnum
    Set EnumValue = CreateclsEnum
    Select Case fsoFilter.GetExtensionName(txtBoxDefaultDBFile)
    Case EnumValue.DBFileExetension(xlam_dext), EnumValue.DBFileExetension(xlsm_dext), _
    EnumValue.DBFileExetension(xls_dext), EnumValue.DBFileExetension(xlsb_dext), EnumValue.DBFileExetension(xlsx_dext)
        'エクセル関連ファイルの時
        '非表示で処理するために、ApplicationオブジェクトとWorkbookオブジェクトを別に定義する
        Dim objExcel As Excel.Application
        Set objExcel = New Excel.Application
        Dim sqlBC As clsSQLStringBuilder
        Set sqlBC = CreateclsSQLStringBuilder
        Dim wkbFilter As Workbook
        'workbookオブジェクトを取得
        Set wkbFilter = objExcel.Workbooks.Open(fsoFilter.BuildPath(txtBoxDefaultDBDirectory, txtBoxDefaultDBFile))
        '存在しないシート名を開いた場合はErr.Number = 9 、インデックス外エラーが発生するので、エラートラップを行う
        Err.Clear
        Dim shtZaikoInfo As Worksheet
        '在庫情報シートのオブジェクトを取得する、このタイミングでシートが存在しない場合はエラーが発生する
        On Error Resume Next
        Set shtZaikoInfo = wkbFilter.Worksheets(INV_CONST.INV_SH_ZAIKO_NAME)
        On Error GoTo 0
        'Err.Numberが0以外の時は処理を中断
        If Err.Number <> 0 Then
            GoTo CloseAndExit
            Exit Sub
        End If
        If shtZaikoInfo.AutoFilterMode = False Then
            '在庫情報シートにフィルターが設定されていない場合
            Dim rngZaikoInfoColumn As Range
            '在庫情報の列名のうち一つを検索し、Rangeオブジェクトを得る
            Set rngZaikoInfoColumn = shtZaikoInfo.Cells.Find(INV_CONST.F_SH_ZAIKO_TEHAI_TEXT)
            If Not rngZaikoInfoColumn Is Nothing Then
                '手配コードの列が見つかった場合
                '手配コードの列を基準にしてオートフィルタを設定する
                rngZaikoInfoColumn.AutoFilter
                'フィルタ設定した状態でブックを保存する
                wkbFilter.Save
            End If
        Else
            'フィルタモードが有効になっている場合
            'フィルタ設定範囲が名前定義として存在しているが、非表示になっているので表示する設定にする
            '名前定義すべてに対してループする
            Dim elmName As Name
            Dim flgSave As Boolean
            '保存フラグをFalseで初期化する
            flgSave = False
            For Each elmName In shtZaikoInfo.Names
                If elmName.Visible = False Then
                    '名前定義が非表示なっていた場合
                    elmName.Visible = True
                    '保存フラグを立てる
                    flgSave = True
                End If
            Next elmName
            '保存フラグの状態を調べる
            If flgSave Then
                '保存フラグが立っていたらブックを保存する
                wkbFilter.Save
            End If
        End If
    End Select
    GoTo CloseAndExit
CloseAndExit:
    Set shtZaikoInfo = Nothing
    If Not wkbFilter Is Nothing Then
        'WorkBookオブジェクトがNothingではない場合
        wkbFilter.Close
        Set wkbFilter = Nothing
    End If
    If Not objExcel Is Nothing Then
        'Excell.ApplicationオブジェクトがNothingではない場合
        objExcel.Quit
        Set objExcel = Nothing
    End If
    Set fsoFilter = Nothing
    Set EnumValue = Nothing
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
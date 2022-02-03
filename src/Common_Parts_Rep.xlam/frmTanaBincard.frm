VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTanaBincard 
   Caption         =   "棚卸BINカードチェック用フォーム"
   ClientHeight    =   8490.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8160
   OleObjectBlob   =   "frmTanaBincard.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTanaBincard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'フォーム内共有変数
Private clsADOfrmBIN As clsADOHandle
Private clsINVDBfrmBIN As clsINVDB
Private clsEnumfrmBIN As clsEnum
Private clsSQLBc As clsSQLStringBuilder
Private objExcelFrmBIN As Excel.Application
Private dicoObjNameToFieldName As Dictionary
Private clsIncrementalfrmBIN As clsIncrementalSerch
'メンバ変数
Private rsfrmBIN As ADODB.Recordset
'------------------------------------------------------------------------
'定数定義
'T_INV_CSVはここ位でしか扱わないので、Privateでも大丈夫
'F_CSV_Status
Private Const CSV_STATUS_BIN_INPUT As Long = &H1    'BINカード残数がNullじゃない
Private Const CSV_STATUS_BIN_DATAOK As Long = &H2   'BINカード残数とデータ残数が一致
Private Const CSV_STATUS_REAL_INPU As Long = &H4    '現品残がNullじゃない
Private Const CSV_STATUS_REAL_DATAOK As Long = &H8  '現品残とデータ残数が一致
'SQL
'棚卸締切日データ取得SQL
'{0}    締切日
'{1}    T_INV_CSV
'{2}    (AfterINWord)
Private Const CSV_SQL_ENDDAY_LIST As String = "SELECT DISTINCT {0} FROM {1} IN""""{2}"
'棚卸チェック用デフォルトデータ取得SQL
'{0}    (selectField As 必須)
'{1}    T_INV_CSV
'{2}    (After IN Word)
'{3}    (TCSVtana? Alias)
'{4}    ロケーション
'{5}    締切日
'{6}    (lstBox_EndDayの選択テキスト)
'{7}    (追加するWhere条件あれば、なければ"")
'{8}    (ORDER BY引数 F_ロケーション ASC ？)
Private Const CSV_SQL_TANA_DEFAULT As String = "SELECT {0} FROM {1} IN """"{2} AS {3} " & vbCrLf & _
"WHERE {4} LIKE ""K%"" AND LEN{4} >= 2 AND {5} = {6} {7}" & vbCrLf & _
"ORDER BY {8}"
Private Sub btnRegistTanaCSVtoDB_Click()
    '最初にCSVファイルを選択してもらう
    Dim strCSVFullPath As String
    'カレントディレクトリをダウンロードディレクトリに変更する
    Call ChCurrentDirW(GetDownloadPath)
    MsgBox "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"
    strCSVFullPath = CStr(Application.GetOpenFilename("CSVファイル,*.csv", 1, "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"))
    If strCSVFullPath = "False" Then
        'キャンセルボタンが押された
        MsgBox "キャンセルしました"
        Exit Sub
    End If
    Dim longAffected As Long
#If DontRemoveZaikoSH Then
    'DLしたファイルを残しておく（テスト環境向け）
    '取得したファイル名を引数にしてDBに登録（拡張子によって処理が分岐されるはず）
    longAffected = clsINVDBfrmBIN.UpsertINVPartsMasterfromZaikoSH(strCSVFullPath, objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc, True)
#Else
    longAffected = clsINVDBfrmBIN.UpsertINVPartsMasterfromZaikoSH(strCSVFullPath, objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc, False)
#End If
End Sub
Private Sub lstBoxEndDay_Click()
    '締切日リスト選択された
    '選択された締切日からデータ取得し、メンバ変数のrsにセットしてやる
    Dim isCollect As Boolean
    isCollect = setDefaultDataToRS(lstBoxEndDay.List(lstBoxEndDay.ListIndex))
    If Not isCollect Then
        MsgBox "棚卸締切日: " & lstBoxEndDay.List(lstBoxEndDay.ListIndex) & " のデータの取得に失敗しました"
        Exit Sub
    End If
End Sub
'フォーム初期化動作
Private Sub UserForm_Initialize()
    'メンバインスタンス変数セット
    If clsADOfrmBIN Is Nothing Then
        Set clsADOfrmBIN = CreateclsADOHandleInstance
    End If
    If clsINVDBfrmBIN Is Nothing Then
        Set clsINVDBfrmBIN = CreateclsINVDB
    End If
    If clsEnumfrmBIN Is Nothing Then
        Set clsEnumfrmBIN = CreateclsEnum
    End If
    If clsSQLBc Is Nothing Then
        Set clsSQLBc = CreateclsSQLStringBuilder
    End If
    If objExcelFrmBIN Is Nothing Then
        Set objExcelFrmBIN = New Excel.Application
    End If
    If dicoObjNameToFieldName Is Nothing Then
        Set dicoObjNameToFieldName = New Dictionary
    End If
    If clsIncrementalfrmBIN Is Nothing Then
        Set clsIncrementalfrmBIN = CreateclsIncrementalSerch
    End If
    If rsfrmBIN Is Nothing Then
        Set rsfrmBIN = New ADODB.Recordset
    End If
    '棚卸締切日リストを設定
    Dim isCollect As Boolean
    isCollect = setEndDayList
    If Not isCollect Then
        MsgBox "棚卸CSVのDBデータ読み込みでエラーが発生しました"
        Unload Me
    End If
End Sub
'''締切日リストを設定する
'''Return Bool  成功したらTrue、それ以外はFalse
Private Function setEndDayList() As Boolean
    On Error GoTo ErrorCatch
''{0}    締切日
''{1}    T_INV_CSV
''{2}    (AfterINWord)
'Private Const CSV_SQL_ENDDAY_LIST As String = "SELECT DISTINCT {0} FROM {1} IN""""{2}"
    Dim dicReplaceEndDay As Dictionary
    Set dicReplaceEndDay = New Dictionary
    dicReplaceEndDay.Add 0, clsEnumfrmBIN.CSVTanafield(F_EndDay_ICS)
    dicReplaceEndDay.Add 1, INV_CONST.T_INV_CSV
    'DBPathをデフォルトへ
    clsADOfrmBIN.SetDBPathandFilenameDefault
    Dim fsoEndDay As FileSystemObject
    Set fsoEndDay = New FileSystemObject
    dicReplaceEndDay.Add 2, clsSQLBc.CreateAfterIN_WordFromSHFullPath(fsoEndDay.BuildPath(clsADOfrmBIN.DBPath, clsADOfrmBIN.DBFileName), clsEnumfrmBIN)
    '置換実行、SQL設定
    clsADOfrmBIN.SQL = clsSQLBc.ReplaceParm(CSV_SQL_ENDDAY_LIST, dicReplaceEndDay)
    Dim isCollect As Boolean
    'SQL実行
    isCollect = clsADOfrmBIN.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        MsgBox "setEndDayList 棚卸CSVのDBデータ読み取りに失敗しました"
        setEndDayList = False
        GoTo CloseAndExit
    End If
    '一旦2次元配列で、タイトル無しの配列を受け取る
    Dim SQL2DimmentionResult() As Variant
    SQL2DimmentionResult = clsADOfrmBIN.RS_Array(True)
    '次に1次元配列に変換したものを受け取る
    Dim SQL1DimmentionList() As Variant
    SQL1DimmentionList = clsSQLBc.SQLResutArrayto1Dimmention(SQL2DimmentionResult)
    'リストボックスに設定してやる
    lstBoxEndDay.Clear
    lstBoxEndDay.List = SQL1DimmentionList
    'Trueを返して終了
    setEndDayList = True
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "setEndDayList code: " & err.Number & " Description: " & err.Description
    setEndDayList = False
    GoTo CloseAndExit
CloseAndExit:
    Set dicReplaceEndDay = Nothing
    Set fsoEndDay = Nothing
    Exit Function
End Function
'''デフォルト(フィルタ掛かる前）のSelect結果をRSに入れる
'''Retrun bool 成功したらTrue、それ以外はfalse
'''args
'''strEndDay        締切日の10文字
Private Function setDefaultDataToRS(strEndDay As String) As Boolean
    On Error GoTo ErrorCatch
    '設定された引数を元にSQLを組み立てる
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "setDefaultDataToRS code: " & err.Number & " Description: " & err.Description
    setDefaultDataToRS = False
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
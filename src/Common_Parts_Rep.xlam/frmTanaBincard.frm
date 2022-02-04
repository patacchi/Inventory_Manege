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
Private dicObjNameToFieldName As Dictionary
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
    '登録処理が終わったら、リストを再構成する
    setEndDayList
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
    If dicObjNameToFieldName Is Nothing Then
        Set dicObjNameToFieldName = New Dictionary
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
    'divObjToFieldを設定
    setDicObjToField
End Sub
Private Sub ClearAllContents(strargExceptControlName As String)
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
'''dicObjToFieldNameの設定を行う
'''key がオブジェクト名、value がテーブルエイリアス付きフィールド名
Private Sub setDicObjToField()
    On Error GoTo ErrorCatch
    If dicObjNameToFieldName Is Nothing Then
        '初期化されていなかったら初期化する
        Set dicObjNameToFieldName = New Dictionary
    End If
    '最初に全消去
    dicObjNameToFieldName.RemoveAll
    dicObjNameToFieldName.Add txtBox_F_CSV_No.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_CSV_No_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_Tana_Local_Text.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Location_Text_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_Tehai_Code.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Tehai_Code_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_DB_Amount.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Stock_Amount_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_BIN_Amount.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Bin_Amount_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_Real_Amount.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Available_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_System_Name.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_System_Name_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_System_Spac.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_System_Spec_ICS), clsEnumfrmBIN)
    '以下は画面表示はしないものの、RSでデータとして保持はするものなので、KeyはDBのフィールド名（テーブルエイリアスプレフィックス無し）、Valueはプレフィックス有りとする
    dicObjNameToFieldName.Add clsEnumfrmBIN.CSVTanafield(F_Status_ICS), clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Status_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add clsEnumfrmBIN.CSVTanafield(F_EndDay_ICS), clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_EndDay_ICS), clsEnumfrmBIN)
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "dicObjNameToFieldName code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''デフォルト(フィルタ掛かる前）のSelect結果をRSに入れる
'''Retrun bool 成功したらTrue、それ以外はfalse
'''args
'''strargEndDay        締切日の10文字
Private Function setDefaultDataToRS(strargEndDay As String) As Boolean
    On Error GoTo ErrorCatch
    '設定された引数を元にSQLを組み立てる
''棚卸チェック用デフォルトデータ取得SQL
''{0}    (selectField As 必須)
''{1}    T_INV_CSV
''{2}    (After IN Word)
''{3}    (TCSVtana? Alias)
''{4}    ロケーション
''{5}    締切日
''{6}    (lstBox_EndDayの選択テキスト)
''{7}    (追加するWhere条件あれば、なければ"")
''{8}    (ORDER BY引数 F_ロケーション ASC ？)
'Private Const CSV_SQL_TANA_DEFAULT As String = "SELECT {0} FROM {1} IN """"{2} AS {3} " & vbCrLf &
    '置換用dic宣言、初期化
    Dim dicReplaceSetDefault As Dictionary
    Set dicReplaceSetDefault = New Dictionary
    'DBPathをデフォルトに
    clsADOfrmBIN.SetDBPathandFilenameDefault
    dicReplaceSetDefault.RemoveAll
    Dim strSelectField As String
    strSelectField = clsSQLBc.GetSELECTfieldListFromDicObjctToFieldName(dicObjNameToFieldName)
    dicReplaceSetDefault.Add 0, strSelectField
    dicReplaceSetDefault.Add 1, INV_CONST.T_INV_CSV
    Dim fsoSetDefault As FileSystemObject
    Set fsoSetDefault = New FileSystemObject
    dicReplaceSetDefault.Add 2, clsSQLBc.CreateAfterIN_WordFromSHFullPath(fsoSetDefault.BuildPath(clsADOfrmBIN.DBPath, clsADOfrmBIN.DBFileName), clsEnumfrmBIN)
    dicReplaceSetDefault.Add 3, clsEnumfrmBIN.SQL_INV_Alias(TanaCSV_Alias_sia)
    dicReplaceSetDefault.Add 4, clsEnumfrmBIN.CSVTanafield(F_Location_Text_ICS)
    dicReplaceSetDefault.Add 5, clsEnumfrmBIN.CSVTanafield(F_EndDay_ICS)
    dicReplaceSetDefault.Add 6, strargEndDay
    '(追加条件があればここで加味する)
    'とりあえずは絞り込みなし
    dicReplaceSetDefault.Add 7, ""
    dicReplaceSetDefault.Add 8, Replace(dicObjNameToFieldName(txtBox_F_CSV_Tana_Local_Text.Name), ".", "_") & " ASC"
    'Replace実行、SQL設定
    clsADOfrmBIN.SQL = clsSQLBc.ReplaceParm(CSV_SQL_TANA_DEFAULT, dicReplaceSetDefault)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "setDefaultDataToRS code: " & err.Number & " Description: " & err.Description
    setDefaultDataToRS = False
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
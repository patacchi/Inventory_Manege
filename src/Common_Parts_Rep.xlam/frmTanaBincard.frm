VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTanaBincard 
   Caption         =   "棚卸BINカードチェック用フォーム"
   ClientHeight    =   9555.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8115
   OleObjectBlob   =   "frmTanaBincard.frx":0000
   ShowModal       =   0   'False
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
Private confrmBIN As ADODB.Connection
Private StopEvents As Boolean
Public strOriginEndDay As String                        '継承元となるEndDayを格納する変数
'------------------------------------------------------------------------
'定数定義
'T_INV_CSVはここ位でしか扱わないので、Privateでも大丈夫
'F_CSV_Status
Private Const CSV_STATUS_BIN_INPUT As Long = &H1    'BINカード残数がNullじゃない
Private Const CSV_STATUS_BIN_DATAOK As Long = &H2   'BINカード残数とデータ残数が一致
Private Const CSV_STATUS_REAL_INPU As Long = &H4    '現品残がNullじゃない
Private Const CSV_STATUS_REAL_DATAOK As Long = &H8  '現品残とデータ残数が一致
'enum
'Status
Private Enum Enum_frmBIN_Status
    BINInput = &H1
    BINDataOK = &H2
    RealInput = &H4
    RealDataOK = &H8
    AllOK = Enum_frmBIN_Status.BINInput Or Enum_frmBIN_Status.BINDataOK Or Enum_frmBIN_Status.RealInput Or Enum_frmBIN_Status.RealDataOK
End Enum
'------------------------------------------------------------------------
'SQL
'棚卸締切日データ取得SQL
'{0}    締切日
'{1}    T_INV_CSV
'{2}    (AfterINWord)
Private Const CSV_SQL_ENDDAY_LIST As String = "SELECT DISTINCT {0} FROM {1} IN""""{2} ORDER BY {0} ASC"
'棚卸チェック用デフォルトデータ取得SQL
'{0}    (selectField As 必須)
'{1}    T_INV_CSV
'{2}    (After IN Word)
'{3}    (TCSVtana? Alias)
'{4}    ロケーション
'{5}    締切日
'{6}    (lstBox_EndDayの選択テキスト)
'{7}    (追加するWhere条件あれば、なければ"")
'{8}    (ORDER BY引数 TCSVTana.ロケーション ASC ？)
Private Const CSV_SQL_TANA_DEFAULT As String = "SELECT {0} FROM {1}  AS {3} IN""""{2}  " & vbCrLf & _
"WHERE {3}.{4} LIKE ""K%"" AND LEN({3}.{4}) >= 2 AND {3}.{5} = ""{6}"" {7}" & vbCrLf & _
"ORDER BY {8}"
'追加条件用
'指定ビットが立っていない物のみ抽出
'NOT ((TCSVTana.F_CSV_Status MOD(2*2)) >= 2)
'{0}    TCSVTana.F_CSV_Status
'{1}    2 落としたいビットだけを上げたLong
Private Const CSV_SQL_BIT_NOT_INCLUDE As String = "NOT (({0} MOD({1}*2)) >= {1})"
'文字列でLike条件
'TCSVTana.ロケーション LIKE ""K%""
'{0}    TCSVTana.ロケーション
'{1}    K 条件文字列、前方一致
Private Const CSV_SQL_WHERE_LIKE As String = "{0} LIKE ""{1}%"""
'------------------------------------------------------------------------------------------------
'BINカード残数と現品残を継承するSQL
'{0}    T_INV_CSV
'{1}    T_Dst
'{2}    ロケーション
'{3}    棚卸締切日
'{4}    (Origin EndDay)
'{5}    T_Orig
'{6}    手配コード
'{7}    F_CSV_BIN_Amount
'{8}    現品残
'{9}    (Dst EndDay)
Private Const CSV_SQL_INHERIT_AMOUNT As String = "UPDATE {0} AS {1} " & vbCrLf & _
"    INNER JOIN (" & vbCrLf & _
"        SELECT * FROM {0} " & vbCrLf & _
"        WHERE {2} LIKE ""K%"" AND LEN({2}) >= 2 AND {3} = ""{4}""" & vbCrLf & _
"        ) AS {5} " & vbCrLf & _
"    ON {1}.{6} = {5}.{6} " & vbCrLf & _
"SET {1}.{7} = {5}.{7},{1}.{8} = {5}.{8} " & vbCrLf & _
"WHERE {1}.{3} = ""{9}"""
'---------------------------------------------------------------------------------------------------------------------
'イベントハンドラ
'フォーム初期化動作
Private Sub UserForm_Initialize()
    ConstRactor
End Sub
'フォーム終了時動作
Private Sub UserForm_Terminate()
    Destractor
End Sub
'DBの内容を既存のCSV(xlsxでも)に追記する
Private Sub btnSaveDBtoCSV_Click()
    If lstBoxEndDay.ListIndex = -1 Then
        MsgBox "棚卸締切日を選択して下さい。"
        Exit Sub
        GoTo CloseAndExit
    End If
    On Error GoTo ErrorCatch
    'カレントをダウンロードディレクトリへ移動
    ChCurrentDirW GetDownloadPath
    MsgBox "情報を追記するCSVファイルを選択して下さい"
    Dim strSaveToFileName As String
    strSaveToFileName = CStr(Application.GetOpenFilename("CSVファイル,*.csv", 1, "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"))
    If strSaveToFileName = "False" Then
        'キャンセルが押された
        MsgBox "キャンセルしました"
        GoTo CloseAndExit
        Exit Sub
    End If
    '実行前に自フォームのRSとConnectionの接続を切断してやる
    'イベント停止
    StopEvents = True
    '全項目消去
    ClearAllContents
    'RSとConnecitonの接続を切断
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        clsADOfrmBIN.RS.ActiveConnection.Close
    End If
    If confrmBIN.State And ObjectStateEnum.adStateOpen Then
        confrmBIN.Close
    End If
    'clsINVDBの初期化
    Dim isCollect As Boolean
    isCollect = clsINVDBfrmBIN.SetShareInsance(objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc)
    If Not isCollect Then
        MsgBox "共有インスタンスの設定でエラーが発生しました。"
        GoTo CloseAndExit
        Exit Sub
    End If
    Dim strBackUpFilePath As String
    strBackUpFilePath = clsINVDBfrmBIN.SetDBDatatoTanaCSV(strSaveToFileName, lstBoxEndDay.List(lstBoxEndDay.ListIndex))
    If strBackUpFilePath = "" Then
        MsgBox "CSVファイル追記中にエラーが発生しました"
        GoTo CloseAndExit
        Exit Sub
    End If
    MsgBox strSaveToFileName & vbCrLf & " ファイルに追記を行い、オリジナルのファイルは次の名前にリネームしました " & vbCrLf & strBackUpFilePath
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "btnSaveDBtoCSV code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
    Exit Sub
CloseAndExit:
    'RS再取得
    setDefaultDatatoRS (lstBoxEndDay.List(lstBoxEndDay.ListIndex))
    'RSよりデータ取得
    getValueFromRS
    'イベント再開
    StopEvents = False
End Sub
'棚卸CSVからDBに登録するボタン
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
    'RSとConnecitonが接続していたら切断する
    'RS
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        clsADOfrmBIN.RS.ActiveConnection.Close
    End If
    'Connection
    If confrmBIN.State And ObjectStateEnum.adStateOpen Then
        confrmBIN.Close
    End If
'#If DontRemoveZaikoSH Then
    'DLしたファイルを残しておく（テスト環境向け）
    'DLしたCSVに書き戻す動作も追加したので、CSVファイルはそのまま残した方がいいかも
    '取得したファイル名を引数にしてDBに登録（拡張子によって処理が分岐されるはず）
    longAffected = clsINVDBfrmBIN.UpsertINVPartsMasterfromZaikoSH(strCSVFullPath, objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc, True)
''#Else
    '以下はファイル削除する場合だけれども・・・
'    longAffected = clsINVDBfrmBIN.UpsertINVPartsMasterfromZaikoSH(strCSVFullPath, objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc, False)
'#End If
    '登録処理が終わったら、もう一度初期設定をし、リストを再構成する
    ConstRactor
End Sub
'''BINカード残数、現品残を前のデータから継承する
Private Sub btnInheritAmount_Click()
    InheritAmount
End Sub
Private Sub lstBoxEndDay_Click()
    '締切日リスト選択された
    '選択された締切日からデータ取得し、メンバ変数のrsにセットしてやる
    Dim isCollect As Boolean
    isCollect = setDefaultDatatoRS(lstBoxEndDay.List(lstBoxEndDay.ListIndex))
    If Not isCollect Then
        MsgBox "棚卸締切日: " & lstBoxEndDay.List(lstBoxEndDay.ListIndex) & " のデータの取得に失敗しました"
        Exit Sub
    End If
    'イベントを停止する
    StopEvents = True
    'RSより取得するデータ全クリア
    ClearAllContents
    'データ総数をラベルにセット
    lbl_TotalAmount.Caption = CStr(clsADOfrmBIN.GetRecordCountFromRS(clsADOfrmBIN.RS))
    'RSから値取得、表示
    getValueFromRS
    'イベントを再開
    StopEvents = False
End Sub
'インクリメンタルサーチのリストくりこ
Private Sub lstBox_IncrementalSerch_Click()
    On Error GoTo ErrorCatch
    'イベント停止
    StopEvents = True
    If clsIncrementalfrmBIN.Incremental_LstBox_Click Then
        'この中に入ってる時点でRSにフィルタがかかってる
        'clsincrementalのイベントも停止する
        clsIncrementalfrmBIN.StopEvent = True
        'AditionalFilterのためにフィルターテキストボックスにインクリメンタルリストの値をセット
        clsIncrementalfrmBIN.txtBoxRef.Text = lstBox_IncrementalSerch.List(lstBox_IncrementalSerch.ListIndex)
'        インクリメンタルリストボックスを非表示､はKeyupとMouseUpに任せた
        'RSのデータ反映
        getValueFromRS True
'        '追加条件設定はここではしなくてもいい
'        AditionalWhereFilter clsIncrementalfrmBIN.txtBoxRef
        'clsIncremantalのイベントも再開してやる
        clsIncrementalfrmBIN.StopEvent = False
    End If
    'イベント再開
    StopEvents = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "lstBox_IncrementalSerch_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'前後移動ボタン
Private Sub btnMoveNextData_Click()
    '次へ
    MoveRecord vbKeyRight
End Sub
Private Sub btnMovePreviosData_Click()
    '前へ
    MoveRecord vbKeyLeft
End Sub
'No Real
Private Sub chkBoxShowNOReal_Click()
    If StopEvents Then
        'イベント停止フラグが立っていた
        Exit Sub
    End If
    AditionalWhereFilter ActiveControl
End Sub
'No BIN
Private Sub chkBoxShowNotBIN_Click()
    If StopEvents Then
        'イベント停止フラグが立っていた
        Exit Sub
    End If
    AditionalWhereFilter ActiveControl
End Sub
'Enterイベント
'棚番フィルターテキストボックスEnter
Private Sub txtBox_Filter_F_CSV_Tana_Local_Text_Enter()
    'イベント停止する
    StopEvents = True
    'clsIncremtental.txtBoxrefにあくちぶコントロールを設定
    Set clsIncrementalfrmBIN.txtBoxRef = txtBox_Filter_F_CSV_Tana_Local_Text
    '追加条件設定してやる
    AditionalWhereFilter clsIncrementalfrmBIN.txtBoxRef
    'フィルタテキストボックスと表示テキストボックスの内容が違うときのみインクリメンタル
    If txtBox_F_CSV_Tana_Local_Text.Text <> clsIncrementalfrmBIN.txtBoxRef.Text Then
        'イベント停止する
        StopEvents = True
        clsIncrementalfrmBIN.Incremental_TextBox_Enter clsIncrementalfrmBIN.txtBoxRef, lstBox_IncrementalSerch
        'RSより値を取得しなおし(フォーカス移動無し)
        getValueFromRS True
    End If
    'イベント再開する
    StopEvents = False
End Sub
Private Sub txtBox_Filter_F_CSV_Tehai_Code_Enter()
    'イベント停止する
    StopEvents = True
    'clsIncrementalのtxtBoxの参照にあくちぶコントロールを設定してやる
    Set clsIncrementalfrmBIN.txtBoxRef = txtBox_Filter_F_CSV_Tehai_Code
    '追加条件設定してやる
    AditionalWhereFilter clsIncrementalfrmBIN.txtBoxRef
    'フィルタテキストボックスと手配コードが違うときのみエンターイベント
    If txtBox_F_CSV_Tehai_Code.Text <> clsIncrementalfrmBIN.txtBoxRef.Text Then
        'イベント停止する
        StopEvents = True
        clsIncrementalfrmBIN.Incremental_TextBox_Enter clsIncrementalfrmBIN.txtBoxRef, lstBox_IncrementalSerch
        'インクリメンタルサーチで一旦値消去されてるので、取得しなおす
        getValueFromRS True
    End If
    'イベント再開する
    StopEvents = False
End Sub
'インクリメンタルリストEnter
Private Sub lstBox_IncrementalSerch_Enter()
    'イベント停止する
    StopEvents = True
    clsIncrementalfrmBIN.Incremental_LstBox_Enter
    'イベント再開する
    StopEvents = False
End Sub
Private Sub btnMoveNextData_Enter()
    If lstBox_IncrementalSerch.ListCount <= 1 Then
        lstBox_IncrementalSerch.Height = 0
    Else
        lstBox_IncrementalSerch.Visible = False
    End If
End Sub
Private Sub btnMovePreviosData_Enter()
    If lstBox_IncrementalSerch.ListCount <= 1 Then
        lstBox_IncrementalSerch.Height = 0
    Else
        lstBox_IncrementalSerch.Visible = False
    End If
End Sub
Private Sub txtBox_F_CSV_BIN_Amount_Enter()
    If lstBox_IncrementalSerch.ListCount <= 1 Then
        lstBox_IncrementalSerch.Height = 0
    Else
        lstBox_IncrementalSerch.Visible = False
    End If
End Sub
Private Sub txtBox_F_CSV_Real_Amount_Enter()
    If lstBox_IncrementalSerch.ListCount <= 1 Then
        lstBox_IncrementalSerch.Height = 0
    Else
        lstBox_IncrementalSerch.Visible = False
    End If
End Sub
'Changeイベント
Private Sub txtBox_F_CSV_BIN_Amount_Change()
    'BINカード残数
    If StopEvents Then
        'イベント停止フラグが立っていたら中止
        Exit Sub
    End If
    '空白で消去（Nullセット)するようにするので、空白でイベントはありえる
    'データチェックとDB登録を実行
    CheckDataAndUpdateDB txtBox_F_CSV_BIN_Amount
End Sub
Private Sub txtBox_F_CSV_Real_Amount_Change()
    '実残数
    If StopEvents Then
        'イベント停止フラグが立っていた
        Exit Sub
    End If
    'データチェックとDB登録を実行
    CheckDataAndUpdateDB txtBox_F_CSV_Real_Amount
End Sub
'Filter_Location
Private Sub txtBox_Filter_F_CSV_Tana_Local_Text_Change()
    If StopEvents Then
        'イベント停止フラグが立っていた
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'まずはUCASEを掛ける
    If txtBox_Filter_F_CSV_Tana_Local_Text.Text <> "" Then
        txtBox_Filter_F_CSV_Tana_Local_Text.Text = UCase(txtBox_Filter_F_CSV_Tana_Local_Text.Text)
    End If
    '条件絞り込み実行
    AditionalWhereFilter clsIncrementalfrmBIN.txtBoxRef
    'clsIncremental.textBoxrefと表示データが違う場合のみインクリメンタル実行
    If txtBox_F_CSV_Tana_Local_Text.Text <> clsIncrementalfrmBIN.txtBoxRef.Text Then
        'イベント停止する
        StopEvents = True
        clsIncrementalfrmBIN.Incremental_TextBox_Change
        'RSから値を取得しなおす、ふぉかーす移動無し
        getValueFromRS True
    End If
    'イベント再開する
    StopEvents = False
End Sub
'Filter_Tehai_Code
Private Sub txtBox_Filter_F_CSV_Tehai_Code_Change()
    If StopEvents Then
        'イベント停止フラグが立っていた
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'Ucase掛ける
    If txtBox_Filter_F_CSV_Tehai_Code.Text <> "" Then
        txtBox_Filter_F_CSV_Tehai_Code.Text = UCase(txtBox_Filter_F_CSV_Tehai_Code.Text)
    End If
    '条件絞り込み実行
    AditionalWhereFilter clsIncrementalfrmBIN.txtBoxRef
    '絞り込みかけた後clsIncremental.txtBoxrefと表示データが違うときのみインクリメンタルサーチ開始
    If txtBox_F_CSV_Tehai_Code.Text <> clsIncrementalfrmBIN.txtBoxRef.Text Then
        'イベント停止する
        StopEvents = True
        'clsIncreMentalのイベントを再開する
        clsIncrementalfrmBIN.StopEvent = False
        clsIncrementalfrmBIN.Incremental_TextBox_Change
        'インクリメンタルサーチで一旦全リスト消去されてるので、値をRSから取得しなおす
        getValueFromRS True
    End If
    'イベント再開する
    StopEvents = False
End Sub
'KeyDownイベント
Private Sub txtBox_F_CSV_BIN_Amount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If StopEvents Then
        'イベント停止フラグ立っている間はキー入力無効にする
        KeyCode = 0
        Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight
        '左右矢印キー
        If chkBoxInputContiue.Value Then
            '連続入力モードだったら
            'レコード移動
            MoveRecord (KeyCode)
            KeyCode = 0
        End If
    End Select
End Sub
Private Sub txtBox_F_CSV_Real_Amount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If StopEvents Then
        'イベント停止フラグ立っている間はキー入力を無効にする
        KeyCode = 0
        Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight
        '左右矢印キー
        If chkBoxInputContiue.Value Then
            '連続入力モードだったら
            'レコード移動
            MoveRecord (KeyCode)
            KeyCode = 0
        End If
    End Select
End Sub
Private Sub btnMoveNextData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If StopEvents Then
        'イベント停止フラグ立ってたら抜ける
        Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight
        '左右矢印キー
        If chkBoxInputContiue.Value Then
            '連続入力モードだったら
            'レコード移動
            MoveRecord (KeyCode)
            KeyCode = 0
        End If
    End Select
End Sub
Private Sub btnMovePreviosData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If StopEvents Then
        'イベント停止フラグ立ってたら抜ける
        Exit Sub
    End If
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight
        '左右矢印キー
        If chkBoxInputContiue.Value Then
            '連続入力モードだったら
            'レコード移動
            MoveRecord (KeyCode)
            KeyCode = 0
        End If
    End Select
End Sub
'KeyUpイベント
Private Sub lstBox_IncrementalSerch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If StopEvents Then
        'イベント停止フラグ立ってたら抜ける
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    'インクリメンタルリストのキーイベント
    clsIncrementalfrmBIN.Incremental_LstBox_Key_UP KeyCode, Shift
    Select Case KeyCode
    Case vbKeyEscape, vbKeyReturn
        'キーがReturnかESCだった時
        'フォーカス移動目的でRSよりデータ取得
        getValueFromRS
    End Select
    'イベント再開
    StopEvents = False
End Sub
'MouseUpイベント
Private Sub lstBox_IncrementalSerch_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If StopEvents Then
        'イベント停止フラグ立ってたら抜ける
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    'インクリメンタルリストマウスアップイベント
    clsIncrementalfrmBIN.Incremental_LstBox_Mouse_UP Button
    Select Case Button
    Case vbMouseLeft
        'マウス左クリックだった時
        'フォーカス移動目的でRSよりデータ取得
        getValueFromRS
    End Select
    'イベント再開
    StopEvents = False
End Sub
'---------------------------------------------------------------------------------------------------------------------
'プロシージャ
Private Sub ConstRactor()
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
    If clsADOfrmBIN.RS Is Nothing Then
        Set clsADOfrmBIN.RS = New ADODB.Recordset
    End If
    If confrmBIN Is Nothing Then
        Set confrmBIN = New ADODB.Connection
    End If
    '棚卸締切日リストを設定
    Dim isCollect As Boolean
    isCollect = setEndDayList
    If Not isCollect Then
        MsgBox "棚卸CSVのDBデータ読み込みでエラーが発生しました"
'        Unload Me
        Exit Sub
    End If
#If DebugDB Then
    MsgBox "DebugDB Enable"
#End If
    'divObjToFieldを設定
    setDicObjToField
    'clsIncrementalSerchコンストラクタ
    clsIncrementalfrmBIN.ConstRuctor Me, dicObjNameToFieldName, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc
    If chkBox_SelectNewData.Value Then
        '起動時最新データ選択フラグが立っていたらリストの一番下を選択する
        lstBoxEndDay.ListIndex = lstBoxEndDay.ListCount - 1
    End If
    Exit Sub
End Sub
'フォーム終了時に実行するプロシージャ
Private Sub Destractor()
    'メンバ変数の解放
    If Not clsADOfrmBIN.RS Is Nothing Then
        If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
            '接続が開いていたら閉じる
            clsADOfrmBIN.RS.Close
        End If
        Set clsADOfrmBIN.RS = Nothing
    End If
    If Not clsADOfrmBIN Is Nothing Then
        clsADOfrmBIN.CloseClassConnection
        Set clsADOfrmBIN = Nothing
    End If
    If Not clsINVDBfrmBIN Is Nothing Then
        Set clsINVDBfrmBIN = Nothing
    End If
    If Not clsEnumfrmBIN Is Nothing Then
        Set clsEnumfrmBIN = Nothing
    End If
    If Not clsSQLBc Is Nothing Then
        Set clsSQLBc = Nothing
    End If
    If Not objExcelFrmBIN Is Nothing Then
        objExcelFrmBIN.Quit
        Set objExcelFrmBIN = Nothing
    End If
    If Not dicObjNameToFieldName Is Nothing Then
        Set dicObjNameToFieldName = Nothing
    End If
    If Not clsIncrementalfrmBIN Is Nothing Then
        Set clsIncrementalfrmBIN = Nothing
    End If
    If Not confrmBIN Is Nothing Then
        If confrmBIN.State And ObjectStateEnum.adStateOpen Then
            'openフラグが立っていたら閉じる
            confrmBIN.Close
        End If
        Set confrmBIN = Nothing
    End If
    Me.Hide
    Unload Me
    Exit Sub
End Sub
'''dicObjToFieldに存在するコントロールの内容を消去していく
'''args
'''strargExceptControlName  オプション、指定されたNameのオブジェクトのは消去しない
Private Sub ClearAllContents(Optional strargExceptControlName As String)
    'イベント停止する
    StopEvents = True
    Dim controlKey As Control
    For Each controlKey In Me.Controls
        If dicObjNameToFieldName.Exists(controlKey.Name) And (strargExceptControlName <> controlKey.Name) Then
            'dicObjtoFieldに存在し、なおかつ引数の除外コントロール名と一致しなかった場合
            Select Case TypeName(controlKey)
            Case "TextBox"
                'テキストボックスだった場合
                controlKey.Text = ""
            Case "Label"
                'ラベルだった場合
                controlKey.Caption = ""
            End Select
        End If
    Next controlKey
    '消去完了したらイベント再開する
    StopEvents = False
End Sub
'''RSより各コントロールへ値をセットする
'''args
'''NotMoveFocus     Trueにセットすると最後のフォーカス移動をしない
Private Sub getValueFromRS(Optional NotMoveFocus As Boolean = False)
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    '最初に全項目クリア
    Select Case True
    Case clsIncrementalfrmBIN.txtBoxRef Is Nothing
        'インクリメンタルサーチのテキストボックス参照がNothingだった場合
        ClearAllContents
    Case Else
        'インクリメンタルサーチのテキストボックス参照が存在していた場合
        'インクリメンタルサーチ中のテキストボックスは消去しない
        ClearAllContents clsIncrementalfrmBIN.txtBoxRef.Name
    End Select
    'txtBox_F_CSV_No
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_No.Name)).Value) Then
'    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_No.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_No.Text = ""
    Else
        txtBox_F_CSV_No.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_No.Name)).Value
    End If
    'txtBox_F_CSV_Tana_Local_Text
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Tana_Local_Text.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_Tana_Local_Text.Text = ""
    Else
        txtBox_F_CSV_Tana_Local_Text.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Tana_Local_Text.Name)).Value
    End If
    'txtBox_F_CSV_Tehai_Code
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Tehai_Code.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_Tehai_Code.Text = ""
    Else
        txtBox_F_CSV_Tehai_Code.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Tehai_Code.Name)).Value
    End If
    'txtBox_F_CSV_DB_Amount
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_DB_Amount.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_DB_Amount.Text = ""
    Else
        txtBox_F_CSV_DB_Amount.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_DB_Amount.Name)).Value
    End If
    'txtBox_F_CSV_BIN_Amount
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_BIN_Amount.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_BIN_Amount.Text = ""
    Else
        txtBox_F_CSV_BIN_Amount.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_BIN_Amount.Name)).Value
    End If
    'txtBox_F_CSV_Real_Amount
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Real_Amount.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_Real_Amount.Text = ""
    Else
        txtBox_F_CSV_Real_Amount.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Real_Amount.Name)).Value
    End If
    'txtBox_F_CSV_System_Name
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_System_Name.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_System_Name.Text = ""
    Else
        txtBox_F_CSV_System_Name.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_System_Name.Name)).Value
    End If
    'txtBox_F_CSV_System_Spac
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_System_Spac.Name)).Value) Then
        'Nullだった場合
        txtBox_F_CSV_System_Spac.Text = ""
    Else
        txtBox_F_CSV_System_Spac.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_System_Spac.Name)).Value
    End If
    'ステータスチェック
    StatusCheck
    'RSよりStatusを取得
    Dim longStatusValue As Long
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value) Then
        'Nullだった場合
        longStatusValue = 0
    Else
        longStatusValue = CLng(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value)
    End If
    'イベント停止
    StopEvents = True
    '現在のレコード数をCurrentRecordラベルに設定
    lbl_CurrentAmount.Caption = CStr(clsADOfrmBIN.GetRecordCountFromRS(clsADOfrmBIN.RS))
    'パーセントラベル更新
    lbl_PerCent.Caption = Format(CSng(lbl_CurrentAmount.Caption) / CSng(lbl_TotalAmount.Caption), "0%")
    If Not NotMoveFocus Then
        'フォーカス移動禁止フラグが立っていなかったら
        'BINカード残数がデータチェックOKではなかったらフォーカスをBINカード残数へ、OKなら現品残へフォーカス移動
        If longStatusValue = Enum_frmBIN_Status.AllOK Then
            '全データOKの時
            '次へボタンにフォーカス
            btnMoveNextData.SetFocus
        ElseIf Not longStatusValue And BINDataOK Then
            'BINカード残数がデータOKではない
            txtBox_F_CSV_BIN_Amount.SetFocus
            '文字を全選択状態にする
            txtBox_F_CSV_BIN_Amount.SelStart = 0
            txtBox_F_CSV_BIN_Amount.SelLength = Len(txtBox_F_CSV_BIN_Amount.Text)
        ElseIf Not longStatusValue And RealDataOK Then
            '現品残がDataOKじゃない時
            'フォーカスを現品残テキストボックスに移動
            txtBox_F_CSV_Real_Amount.SetFocus
            '文字を全選択状態にする
            txtBox_F_CSV_Real_Amount.SelStart = 0
            txtBox_F_CSV_Real_Amount.SelLength = txtBox_F_CSV_Real_Amount.TextLength
        End If
    End If
    'イベント再開
    StopEvents = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "getValueFromRS code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
Private Sub StatusCheck()
    On Error GoTo ErrorCatch
    'まずRSからStatusの数値を受け取る
    Dim longStatusValue As Long
    If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value) Then
        'Nullだった場合
        longStatusValue = 0
    Else
        longStatusValue = CLng(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value)
    End If
    'Statusに基づきフラグチェックし、表示・非表示を決める
    'Bin Input
    If longStatusValue And Enum_frmBIN_Status.BINInput Then
        lbl_BINcard_Input.Visible = True
        'BIN入力されていたらとりあえず白にしてやる
        txtBox_F_CSV_BIN_Amount.BackColor = &H80000005
    Else
        lbl_BINcard_Input.Visible = False
        '未入力なら水色にしてやる
        txtBox_F_CSV_BIN_Amount.BackColor = &HFFFFC0
    End If
    'BIN Data OK
    If longStatusValue And Enum_frmBIN_Status.BINDataOK Then
        lbl_BINcard_DataOK.Visible = True
        'OKなら白で確定
        txtBox_F_CSV_BIN_Amount.BackColor = &H80000005
    Else
        lbl_BINcard_DataOK.Visible = False
        'Bin NGなら薄い黄色にしてやる
        If longStatusValue And BINInput Then
            'BIN InputがOKの時だけ変更する
            txtBox_F_CSV_BIN_Amount.BackColor = &H80FFFF
        End If
    End If
    'Real Input
    If longStatusValue And Enum_frmBIN_Status.RealInput Then
        lbl_RealAmount_Input.Visible = True
        'とりあえず入力されたら一旦白に
        txtBox_F_CSV_Real_Amount.BackColor = &H80000005
    Else
        lbl_RealAmount_Input.Visible = False
        '未入力なら水色に
        txtBox_F_CSV_Real_Amount.BackColor = &HFFFFC0
    End If
    'Real Data OK
    If longStatusValue And Enum_frmBIN_Status.RealDataOK Then
        lbl_RealAmount_DataOK.Visible = True
        'データOKなら白確定
        txtBox_F_CSV_Real_Amount.BackColor = &H80000005
    Else
        lbl_RealAmount_DataOK.Visible = False
        If longStatusValue And RealInput Then
            'Real NGなら薄い黄色に
            'Real InputOKの時だけ変更する
            txtBox_F_CSV_Real_Amount.BackColor = &H80FFFF
        End If
    End If
    'AllOK
    If longStatusValue = Enum_frmBIN_Status.AllOK Then
        lbl_AllData_OK.Visible = True
        lbl_Data_NG.Visible = False
    Else
        lbl_AllData_OK.Visible = False
        lbl_Data_NG.Visible = True
    End If
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "StatusCheck code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
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
    DebugMsgWithTime "setEndDayList code: " & Err.Number & " Description: " & Err.Description
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
    '棚番ローカルテキストフィルタ
    dicObjNameToFieldName.Add txtBox_Filter_F_CSV_Tana_Local_Text.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Location_Text_ICS), clsEnumfrmBIN)
    dicObjNameToFieldName.Add txtBox_F_CSV_Tehai_Code.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Tehai_Code_ICS), clsEnumfrmBIN)
    '手配コードフィルタ
    dicObjNameToFieldName.Add txtBox_Filter_F_CSV_Tehai_Code.Name, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Tehai_Code_ICS), clsEnumfrmBIN)
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
    DebugMsgWithTime "dicObjNameToFieldName code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''デフォルト(フィルタ掛かる前）のSelect結果をRSに入れる
'''Retrun bool 成功したらTrue、それ以外はfalse
'''args
'''strargEndDay        締切日の10文字
'''strargAditionalWhere 追加のWhere条件をStringで指定
Private Function setDefaultDatatoRS(strargEndDay As String, Optional strargAditionalWhere As String) As Boolean
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
    dicReplaceSetDefault.Add 7, strargAditionalWhere
    dicReplaceSetDefault.Add 8, dicObjNameToFieldName(txtBox_F_CSV_Tana_Local_Text.Name) & " ASC"
    'Replace実行、SQL設定
    clsADOfrmBIN.SQL = clsSQLBc.ReplaceParm(CSV_SQL_TANA_DEFAULT, dicReplaceSetDefault)
    'クラスで接続していたら多重接続になる可能性があるので、明示的に切断してやる
    clsADOfrmBIN.CloseClassConnection
    'SQL組み立て完了したので、データを取り込むRSのプロパティを設定していく
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateConnecting Then
        '接続されていたら一旦切断する
        clsADOfrmBIN.RS.Close
    End If
    If confrmBIN.State And ObjectStateEnum.adStateOpen Then
        'connectionが開いていたら接続を閉じる
        confrmBIN.Close
    End If
    'Connectionオブジェクトの設定を行う
    confrmBIN.ConnectionString = clsADOfrmBIN.CreateConnectionString(clsADOfrmBIN.DBPath, clsADOfrmBIN.DBFileName)
    confrmBIN.Mode = adModeReadWrite Or adModeShareDenyNone
    'レコードセットのSourceとしてclADOのSQLを設定する
    clsADOfrmBIN.RS.Source = clsADOfrmBIN.SQL
    '即時更新モードの状態によって処理を分岐(即時更新かバッチ更新か)
    Select Case chkBoxUpdateASAP
    Case True
        '即時更新有効の場合
        If confrmBIN.State And ObjectStateEnum.adStateOpen Then
            '接続が開いていたら一旦閉じる
            confrmBIN.Close
        End If
        confrmBIN.CursorLocation = adUseServer
        '接続をオープン
        confrmBIN.Open
        'rsのActiveConnectionにメンバ変数のConnectionを設定
        Set clsADOfrmBIN.RS.ActiveConnection = confrmBIN
        'レコードセットを排他的ロックで開く(CursorLocationがCliantの時は共有的ロックでしか開けない)
        clsADOfrmBIN.RS.LockType = adLockPessimistic
        '動的カーソルにし、ほかの人が行った変更は見れるが、追加分は見れないカーソルタイプにする
        'CursorLocationがServerになるので、RecordCountは使えなくなる
        clsADOfrmBIN.RS.CursorLocation = adUseServer
        clsADOfrmBIN.RS.CursorType = adOpenDynamic
        clsADOfrmBIN.RS.Open , , , , CommandTypeEnum.adCmdText
        '登録ボタンを無効に
        btnDoUpdate.Enabled = False
    Case False
        'バッチ更新モード
        If confrmBIN.State And ObjectStateEnum.adStateOpen Then
            '接続が開いていたら一旦閉じる
            confrmBIN.Close
        End If
        confrmBIN.CursorLocation = adUseClient
        '接続を開く
        confrmBIN.Open
        'rsのActiveConnectionにメンバ変数を割り当てる
        Set clsADOfrmBIN.RS.ActiveConnection = confrmBIN
        'ロックタイプをバッチ更新モードにする
        clsADOfrmBIN.RS.LockType = adLockBatchOptimistic
        '動的カーソルにし、他の人が行った変更を見れるカーソルタイプにする
        clsADOfrmBIN.RS.CursorLocation = adUseClient
        'CursorLocationがCliantの時はカーソルタイプはスタティックかFowerdOnlyしか選べない・・のかな？
        clsADOfrmBIN.RS.CursorType = adOpenStatic
        clsADOfrmBIN.RS.Open , , , , CommandTypeEnum.adCmdText
        '登録ボタンを有効に
        btnDoUpdate.Enabled = True
    End Select
    'FilterとしてadFilterFetchedRecordsをセットしてやり、フェッチ済みの行だけにする
    clsADOfrmBIN.RS.Filter = adFilterFetchedRecords
    'ここでBOF、EOFが共にTrueになってたら取得失敗している
    If clsADOfrmBIN.RS.BOF And clsADOfrmBIN.RS.EOF Then
        '取得失敗(0件)
        setDefaultDatatoRS = False
        GoTo CloseAndExit
    Else
        '取得成功
        'フィルタ解除し、最初のレコードに移動
        clsADOfrmBIN.RS.Filter = adFilterNone
        clsADOfrmBIN.RS.MoveFirst
        setDefaultDatatoRS = True
        GoTo CloseAndExit
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "setDefaultDataToRS code: " & Err.Number & " Description: " & Err.Description
    setDefaultDatatoRS = False
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
'''レコード移動
'''args
'''argKeyCode   KeyCode →なら次へ、←なら前へ
Private Sub MoveRecord(argKeyCode As Integer)
    On Error GoTo ErrorCatch
    If clsADOfrmBIN.RS.State = ObjectStateEnum.adStateClosed Then
        'RSが閉じていたら何もせずに抜ける
        DebugMsgWithTime "MoveRecord : RS not open"
        GoTo CloseAndExit
        Exit Sub
    End If
    Do While clsADOfrmBIN.RS.State And (ObjectStateEnum.adStateConnecting Or ObjectStateEnum.adStateExecuting Or ObjectStateEnum.adStateFetching)
        'RSで作業中は待機する
        DebugMsgWithTime "MoveRecord : RS Busy.wait 300 millisec"
        Sleep 300
    Loop
    Select Case argKeyCode
    Case vbKeyRight
        '右の場合、次へ
        'とりあえずMoveNextする
        clsADOfrmBIN.RS.MoveNext
        'EOFの状態を確認
        If clsADOfrmBIN.RS.EOF Then
            '移動前が最終レコードだった場合
            MsgBox "現在のレコードが最終レコードです"
            clsADOfrmBIN.RS.MovePrevious
            GoTo CloseAndExit
        End If
    Case vbKeyLeft
        '左の場合、前へ
        'とりあえずMovePreviousする
        clsADOfrmBIN.RS.MovePrevious
        'BOFの状態を確認
        If clsADOfrmBIN.RS.BOF Then
            '移動前が最初のレコードだった
            MsgBox "現在のレコードが先頭レコードです"
            clsADOfrmBIN.RS.MoveNext
            GoTo CloseAndExit
        End If
    End Select
    'レコード移動したので、RSから値を取得する
    'イベント停止
    StopEvents = True
    getValueFromRS
    'イベント再開
    StopEvents = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "MoveRecord code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''入力された数値をチェックし、RS(DB)をUpdateする
'''TextBoxのChangeイベントから呼ばれる前提
'''args
'''argTxtBox    イベント発生させたコントロールの参照
Private Sub CheckDataAndUpdateDB(ByRef argTxtBox As MSForms.TextBox)
    On Error GoTo ErrorCatch
    If argTxtBox Is Nothing Then
        DebugMsgWithTime "CheckDataAndUpdateDB : Control name is empty"
        GoTo CloseAndExit
    End If
    'イベント停止する
    StopEvents = True
    Select Case argTxtBox.Text
    Case ""
        '空白だった場合
        'chkBoxNoConfirmDeleteの状態によって問い合わせを実施、結果が通れば強制削除モードでUpdateメソッドを呼び出す
        If Not chkBoxNoConfirmatDelete.Value Then
            '削除確認ありの場合
            Dim lonMsgBoxMSG As Long
            lonMsgBoxMSG = MsgBox("テキストボックスが空白になったので、該当のデータを未入力にします。よろしいですか？", vbYesNo)
            If lonMsgBoxMSG = vbNo Then
                '確認したらNoだった
                MsgBox "削除処理を中断します"
                '該当テキストボックスの値をRSに保存されている値に復元する
                If Not IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value) Then
                    argTxtBox.Text = clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value
                End If
                Exit Sub
            End If
        End If
        UpdateSpecificField argTxtBox, True
    Case Else
        '通常動作はこっち
        If Not IsNumeric(argTxtBox.Text) Then
            MsgBox "入力された文字が数値として認識できませんでした"
            DebugMsgWithTime "CheckDataAndUpdateDB : cant cast number txtboxname: " & argTxtBox.Name
            GoTo CloseAndExit
        End If
        If clsADOfrmBIN.RS.State = ObjectStateEnum.adStateClosed Then
            'RSが閉じていたら何もせずに抜ける
            DebugMsgWithTime "CheckDataAndUpdateDB : RS not open."
            GoTo CloseAndExit
        End If
        'テキストボックスに入力された文字とRSの数値が同じだったら何もしない
        If clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value = CDbl(argTxtBox.Text) Then
            GoTo CloseAndExit
        End If
        'フィールド指定でRSに登録実行
        UpdateSpecificField argTxtBox
    End Select
    'フラグチェックと設定
    ChekStatusAndSetFlag
    'RSのフラグ情報をもとにラベルを更新する
    StatusCheck
    '即時更新が有効ならここでUpdateをする
    If chkBoxUpdateASAP.Value Then
        Dim isCollect As Boolean
        isCollect = UpdateDBfromRS
        If Not isCollect Then
            MsgBox "DBへの登録時にエラーが発生しました"
            GoTo CloseAndExit
        End If
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "CheckDataAndUpdateDB code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'''指定されたコントロール名に対応するRSにデータを登録する
'''args
'''argTxtBox    操作対象のコントロールの参照
'''Optional ForceSetNull    Trueにセットすると無条件で指定フィールドにNullをセットする（消去モード）
Private Sub UpdateSpecificField(ByRef argTxtBox As MSForms.TextBox, Optional ForceSetNull As Boolean = False)
    On Error GoTo ErrorCatch
    Select Case ForceSetNull
    Case True
        '強制Nullセットモードはこっち
        '対応するRSにNullをセットする
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value = Null
        GoTo CloseAndExit
    Case False
        '通常動作はこっち
        If argTxtBox.Text = "" Then
            '対象のテキストが空白だったら何もせずに抜ける
            GoTo CloseAndExit
            Exit Sub
        End If
        'RSの値と違っていたら対応するRSに値をセットする
        If IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value) Or clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value <> CDbl(argTxtBox.Text) Then
            'ValueがNullか、テキストボックスの数値と違っていた場合
            clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, argTxtBox.Name)).Value = _
            CDbl(argTxtBox.Text)
        End If
        GoTo CloseAndExit
    End Select
ErrorCatch:
    DebugMsgWithTime "UpdateSpecificField code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'RS上の数値を評価し、フラグの上げ下げを行い、RSに登録する
'RSにテキストボックスの数値を反映させた後に実行する(DB登録エラーのリスクを減らしたい)
Private Sub ChekStatusAndSetFlag()
    On Error GoTo ErrorCatch
'    If clsADOfrmBIN.RS.EditMode = adEditNone Then
'        'RSに変更がないときは何もしない
'        GoTo CloseAndExit
'    End If
    'BinCard
    'Select Case の Caseはショートサーキットになることを利用
    Select Case True
    Case IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_BIN_Amount.Name)).Value)
        'Nullだった場合
        'Bin の Input と DataOKフラグをまとめて落とす
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value And Not (Enum_frmBIN_Status.BINDataOK Or Enum_frmBIN_Status.BINInput)
    Case Not IsNumeric(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_BIN_Amount.Name)).Value)
        '数値として認識されない場合
        'Bin の Input と DataOKフラグをまとめて落とす
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value And Not (Enum_frmBIN_Status.BINDataOK Or Enum_frmBIN_Status.BINInput)
    Case CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_BIN_Amount.Name)).Value) = _
        CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_DB_Amount.Name)).Value)
        'RS上のBINカード残数とシート残数が一致した場合
        'BIN DataOKフラグとBin Inputフラグを両方立てる(入力してないとOKにならない前提)
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value Or Enum_frmBIN_Status.BINDataOK Or Enum_frmBIN_Status.BINInput
    Case CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_BIN_Amount.Name)).Value) <> _
        CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_DB_Amount.Name)).Value)
        'BIN残数とシステム残数が一致しなかった場合
        'Bin Inputフラグを立てて、BI OKフラグを落とす
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        (clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value Or Enum_frmBIN_Status.BINInput) And Not Enum_frmBIN_Status.BINDataOK
    End Select
    'RealAmount
    'Select Case の Caseはショートサーキットになることを利用
    Select Case True
    Case IsNull(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Real_Amount.Name)).Value)
        'Nullだった場合
        'Real の input と DataOK両方まとめて落とす
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value And Not (Enum_frmBIN_Status.RealDataOK Or Enum_frmBIN_Status.RealInput)
    Case Not IsNumeric(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Real_Amount.Name)).Value)
        '数値として認識されない場合
        'Real の input と DataOK両方まとめて落とす
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value And Not (Enum_frmBIN_Status.RealDataOK Or Enum_frmBIN_Status.RealInput)
    Case CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Real_Amount.Name)).Value) = _
        CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_DB_Amount.Name)).Value)
        'RS上のBINカード残数とシート残数が一致した場合
        'Real DataOKフラグとReal Inputフラグを両方立てる(入力してないとOKにならない前提)
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value Or Enum_frmBIN_Status.RealDataOK Or Enum_frmBIN_Status.RealInput
    Case CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_Real_Amount.Name)).Value) <> _
        CDbl(clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, txtBox_F_CSV_DB_Amount.Name)).Value)
        '現品残とシステム上の残数が一致しなかった場合
        'Real Inputフラグを立てて、Real OKフラグを落とす
        clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value = _
        (clsADOfrmBIN.RS.Fields(clsSQLBc.RepDotField(dicObjNameToFieldName, clsEnumfrmBIN.CSVTanafield(F_Status_ICS))).Value Or Enum_frmBIN_Status.RealInput) And Not RealDataOK
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "ChekStatusAndSetFlag code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''RSの内容をDBにUpdateする
'''Return bool 成功したらTrue、それ以外はFalse
Private Function UpdateDBfromRS() As Boolean
    On Error GoTo ErrorCatch
    If clsADOfrmBIN.RS.EditMode = adEditNone Then
        '変更がなかった場合
        '何もせずに抜ける
        DebugMsgWithTime "UpdateDBfromRS : No data changed. Do nothing"
        UpdateDBfromRS = True
        GoTo CloseAndExit
        Exit Function
    End If
    '接続状況を調べ、接続してなかったら接続する
    If Not clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        'RSが未接続だった
        '更にConnectionオブジェクトの接続状況も調べる
        If Not confrmBIN.State And ObjectStateEnum.adStateOpen Then
            'Connectionが未接続だった
            confrmBIN.Open
        End If
        Set clsADOfrmBIN.RS.ActiveConnection = confrmBIN
    End If
    'Update実行する
    clsADOfrmBIN.RS.Update
    If chkBoxUpdateASAP.Value Then
        '即時更新モードの時は、更新完了するまで待機する
        Do Until clsADOfrmBIN.RS.EditMode = adEditNone
            DebugMsgWithTime "UpdateDBfromRS : RS is busy.wait 100 millisec"
            Sleep 100
        Loop
    End If
    'コマンド実行完了まで待機する
    Do While clsADOfrmBIN.RS.State And (ObjectStateEnum.adStateConnecting Or ObjectStateEnum.adStateExecuting Or ObjectStateEnum.adStateFetching)
        DebugMsgWithTime "UpdateDBfromRS : RS is busy.wait 100 millisec"
        Sleep 100
    Loop
    UpdateDBfromRS = True
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "UpdateDBfromRS code: " & Err.Number & " Description: " & Err.Description
    UpdateDBfromRS = False
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
'絞り込みを行う
'''args
'''argSouceCtrl     呼び出し元のコントロールの参照
Private Sub AditionalWhereFilter(ByRef argSouceCtrl As Control)
    On Error GoTo ErrorCatch
    '棚卸締切日選択されてなかったらメッセージ出して抜ける
    If lstBoxEndDay.ListIndex = -1 Then
        '棚卸締切日リストが選択されてなかった
        MsgBox "棚卸締切日が選択されていません。棚卸締切日を選択して下さい。"
        GoTo CloseAndExit
    End If
    If argSouceCtrl Is Nothing Then
        DebugMsgWithTime "AditionalWhereFilter : arg ctrl Nothing"
        GoTo CloseAndExit
    End If
    'インクリメンタルリストを非表示にする
    If lstBox_IncrementalSerch.ListCount >= 2 Then
        lstBox_IncrementalSerch.Visible = False
    End If
    '各コントロールの状態に応じて、追加Where条件を組み立てる
    Dim strarrAddWhere() As String
    ReDim strarrAddWhere(0)
    '0番目には前の条件と繋げるための AND 用に空条件を入れる
    strarrAddWhere(0) = " AND 1=1"
    '以下、それぞれ適切な埋め込み定数を利用し、条件文を組み立てる
    Dim dicReplaceAddWhere As Dictionary
    Set dicReplaceAddWhere = New Dictionary
    'no Bin
    If chkBoxShowNotBIN.Value Then
        '条件配列を拡張
        ReDim Preserve strarrAddWhere(UBound(strarrAddWhere) + 1)
        dicReplaceAddWhere.RemoveAll
        dicReplaceAddWhere.Add 0, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Status_ICS), clsEnumfrmBIN)
        dicReplaceAddWhere.Add 1, BINDataOK
        strarrAddWhere(UBound(strarrAddWhere)) = clsSQLBc.ReplaceParm(CSV_SQL_BIT_NOT_INCLUDE, dicReplaceAddWhere)
    End If
    'no Real
    If chkBoxShowNOReal.Value Then
        '条件配列を拡張
        ReDim Preserve strarrAddWhere(UBound(strarrAddWhere) + 1)
        dicReplaceAddWhere.RemoveAll
        dicReplaceAddWhere.Add 0, clsSQLBc.ReturnTableAliasPlusedFieldName(TanaCSV_Alias_sia, clsEnumfrmBIN.CSVTanafield(F_Status_ICS), clsEnumfrmBIN)
        dicReplaceAddWhere.Add 1, RealDataOK
        strarrAddWhere(UBound(strarrAddWhere)) = clsSQLBc.ReplaceParm(CSV_SQL_BIT_NOT_INCLUDE, dicReplaceAddWhere)
    End If
    'Location
    If txtBox_Filter_F_CSV_Tana_Local_Text.Text <> "" Then
        '条件配列を拡張
        ReDim Preserve strarrAddWhere(UBound(strarrAddWhere) + 1)
        dicReplaceAddWhere.RemoveAll
        dicReplaceAddWhere.Add 0, dicObjNameToFieldName(txtBox_F_CSV_Tana_Local_Text.Name)
        dicReplaceAddWhere.Add 1, txtBox_Filter_F_CSV_Tana_Local_Text.Text
        strarrAddWhere(UBound(strarrAddWhere)) = clsSQLBc.ReplaceParm(CSV_SQL_WHERE_LIKE, dicReplaceAddWhere)
    End If
    'TehaiCode
    If txtBox_Filter_F_CSV_Tehai_Code.Text <> "" Then
        '条件配列を拡張
        ReDim Preserve strarrAddWhere(UBound(strarrAddWhere) + 1)
        dicReplaceAddWhere.RemoveAll
        dicReplaceAddWhere.Add 0, dicObjNameToFieldName(txtBox_F_CSV_Tehai_Code.Name)
        dicReplaceAddWhere.Add 1, txtBox_Filter_F_CSV_Tehai_Code.Text
        strarrAddWhere(UBound(strarrAddWhere)) = clsSQLBc.ReplaceParm(CSV_SQL_WHERE_LIKE, dicReplaceAddWhere)
    End If
    '一旦RSのフィルターを解除する
    clsADOfrmBIN.RS.Filter = adFilterNone
    '指定した条件で条件を指定し、大元のデータを更新してやる
    Dim isCollect As Boolean
    isCollect = setDefaultDatatoRS(lstBoxEndDay.List(lstBoxEndDay.ListIndex), Join(strarrAddWhere, " AND "))
    If Not isCollect Then
        MsgBox "指定条件で適合するレコードがありませんでした。絞り込み条件を元に戻し、データ再取得します。"
        'イベント停止
        StopEvents = True
        '条件設定コントロールの状態を出来るだけ元に戻す
        '呼び出し元のコントロールの種類により処理を分岐
        Select Case TypeName(argSouceCtrl)
        Case "TextBox"
            'テキストボックスだった
            '元々の長さのLen -1 を左から取得
            If argSouceCtrl.TextLength > 1 Then
                'テキストボックスの文字数が1より大きい場合にのみ実行
                argSouceCtrl.Text = Mid(argSouceCtrl.Text, 1, Len(argSouceCtrl.Text) - 1)
            Else
                '文字数が1以下だった場合は、該当テキストボックスの文字を消去する
                argSouceCtrl.Text = ""
            End If
        Case "CheckBox"
            'チェックボックス
            'notで反転する
            argSouceCtrl.Value = Not (argSouceCtrl.Value)
        End Select
'        'イベント再開
'        StopEvents = False
'        '追加条件なしでデータ再取得
'        isCollect = setDefaultDataToRS(lstBoxEndDay.List(lstBoxEndDay.ListIndex))
'        If Not isCollect Then
'            MsgBox "絞り込み条件なしのデータ取得に失敗しました。Excelファイルを開きなおしてもう一度試して下さい。"
'            Unload Me
'            GoTo CloseAndExit
'        End If
'        'イベント停止
'        StopEvents = True
'        'RSから値取得、表示
'        getValueFromRS
        'ダメだった条件を戻してるので、もう1回自信を再帰呼び出しする
        AditionalWhereFilter argSouceCtrl
'        'フォーカス移動無しでRSからデータ取得
'        getValueFromRS True
        'イベント再開
        StopEvents = False
        GoTo CloseAndExit
    End If
    '成功したっぽい
    'イベント停止
    StopEvents = True
    'RSより取得するデータ全クリア
    ClearAllContents argSouceCtrl.Name
    'RSから値取得、表示
    getValueFromRS True
    'イベント再開
    StopEvents = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "AditionalWhereFilter code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set dicReplaceAddWhere = Nothing
    Exit Sub
End Sub
'''BINカード残数、現品残を継承元データから継承する
Private Sub InheritAmount()
    On Error GoTo ErrorCatch
    If lstBoxEndDay.ListCount < 2 Then
        '締切日データが2個未満だったら抜ける
        MsgBox "棚卸締切日のデータが1種類以下です。継承元のデータと継承先のデータを登録してから実行してください"
        GoTo CloseAndExit
    End If
    If lstBoxEndDay.ListIndex = -1 Then
        'リスト選択されていなかったら抜ける
        MsgBox "継承先(新しいデータ)を選択してから実行してください"
        GoTo CloseAndExit
    End If
    '実行していいかどうか問い合わせ
    Dim longMsgBoxReturn As Long
    longMsgBoxReturn = MsgBox("実行すると対象の " & lstBoxEndDay.List(lstBoxEndDay.ListIndex) & " のデータは上書きされますが、よろしいですか？", vbYesNo)
    If longMsgBoxReturn = vbNo Then
        '問い合わせでNoと言われた
        MsgBox "キャンセルしました"
        GoTo CloseAndExit
    End If
    'OriginEndDayメンバ変数初期化
    frmTanaBincard.strOriginEndDay = ""
    'frmSelectOriginEndDay Load
    Load frmSelectOriginEndDay
    '継承先選択フォームの設定、表示
    Dim longEndDayRowCount As Long
    'lstBoxEndDayループ
    For longEndDayRowCount = LBound(lstBoxEndDay.List) To UBound(lstBoxEndDay.List)
        If Not longEndDayRowCount = lstBoxEndDay.ListIndex Then
            '現在選択されている項目と違った場合にOriginEndDayのコンボボックスに追加する
            frmSelectOriginEndDay.cmbBoxOriginEndDay.AddItem lstBoxEndDay.List(longEndDayRowCount)
        End If
    Next longEndDayRowCount
    'SelectOriginEndDayフォームモーダル表示、結果確定したら勝手にUnloadして戻ってくるはず
    frmSelectOriginEndDay.Show
    If frmTanaBincard.strOriginEndDay = "" Then
        'メンバ変数が空文字だったら抜ける
        MsgBox "継承元締切日の選択に失敗しました"
        GoTo CloseAndExit
    End If
    'SQLの組立に入る
''{0}    T_INV_CSV
''{1}    T_Dst
''{2}    ロケーション
''{3}    棚卸締切日
''{4}    (Origin EndDay)
''{5}    T_Orig
''{6}    手配コード
''{7}    F_CSV_BIN_Amount
''{8}    現品残
''{9}    (Dst EndDay)
'Private Const CSV_SQL_INHERIT_AMOUNT As String = "UPDATE {0} AS {1} " & vbCrLf
    Dim dicReplaceInherit As Dictionary
    Set dicReplaceInherit = New Dictionary
    dicReplaceInherit.RemoveAll
    dicReplaceInherit.Add 0, INV_CONST.T_INV_CSV
    dicReplaceInherit.Add 1, clsEnumfrmBIN.SQL_INV_Alias(DstTable_sia)
    dicReplaceInherit.Add 2, clsEnumfrmBIN.CSVTanafield(F_Location_Text_ICS)
    dicReplaceInherit.Add 3, clsEnumfrmBIN.CSVTanafield(F_EndDay_ICS)
    dicReplaceInherit.Add 4, frmTanaBincard.strOriginEndDay
    dicReplaceInherit.Add 5, clsEnumfrmBIN.SQL_INV_Alias(OriginTable_sia)
    dicReplaceInherit.Add 6, clsEnumfrmBIN.CSVTanafield(F_Tehai_Code_ICS)
    dicReplaceInherit.Add 7, clsEnumfrmBIN.CSVTanafield(F_Bin_Amount_ICS)
    dicReplaceInherit.Add 8, clsEnumfrmBIN.CSVTanafield(F_Available_ICS)
    dicReplaceInherit.Add 9, lstBoxEndDay.List(lstBoxEndDay.ListIndex)
    'イベント停止する
    StopEvents = True
    'フォーム共有RSがOpenしてたらCloseする
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        clsADOfrmBIN.RS.Close
    End If
    'clsAdoを単独指定する
    Dim clsAdoInherit As clsADOHandle
    Set clsAdoInherit = CreateclsADOHandleInstance
    'DBPath,DBFilenameをデフォルトへ
    clsAdoInherit.SetDBPathandFilenameDefault
    'Replace実行、SQL設定
    clsAdoInherit.SQL = clsSQLBc.ReplaceParm(CSV_SQL_INHERIT_AMOUNT, dicReplaceInherit)
    Dim isCollect As Boolean
    'Writeフラグ上げる
    clsAdoInherit.ConnectMode = clsAdoInherit.ConnectMode Or adModeWrite
    'SQL実行
    isCollect = clsAdoInherit.Do_SQL_with_NO_Transaction
    'Witeフラグ下げる
    clsAdoInherit.ConnectMode = clsAdoInherit.ConnectMode And Not adModeWrite
    'clsADO切断
    clsAdoInherit.CloseClassConnection
    If Not isCollect Then
        MsgBox "継承SQL実行する際にエラーが発生しました。"
        GoTo CloseAndExit
    End If
    'ここまででDB更新は完了しているが、Statusは自動反映されないのでEndDayリストのClickイベント発生させた後StatusSetしてやる
    '現在のlstboxEndDayのListIndexを退避
    Dim longOldListIndex As Long
    longOldListIndex = lstBoxEndDay.ListIndex
    '一旦EndDayの選択を解除
    lstBoxEndDay.ListIndex = -1
    'イベント再開
    StopEvents = False
    'lstBoxEndDayのLisindexを復元(Clickイベント発生してデータ取りに行くはず)
    lstBoxEndDay.ListIndex = longOldListIndex
    '全レコードのStatus更新
    SetRsAllStatus
    MsgBox "継承動作終了しました"
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "InheritAmount code: " & Err.Number & " Descriptoin: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsAdoInherit Is Nothing Then
        clsAdoInherit.CloseClassConnection
        Set clsAdoInherit = Nothing
    End If
    Exit Sub
End Sub
'''DBのF_CSV_Statusを更新する
Private Sub SetRsAllStatus()
    On Error GoTo ErrorCatch
    If clsADOfrmBIN.RS.State = ObjectStateEnum.adStateClosed Then
        '接続が閉じていたら抜ける
        DebugMsgWithTime "SetRsAllStatus : RS is closed"
        GoTo CloseAndExit
    End If
    'イベント停止する
    StopEvents = True
    'RS全レコードループ
    clsADOfrmBIN.RS.MoveFirst
    Do
        'ステータスセットプロシージャ
        ChekStatusAndSetFlag
        '今回のループのRSを確定
        clsADOfrmBIN.RS.Update
        '次のレコードへ
        clsADOfrmBIN.RS.MoveNext
    Loop While Not clsADOfrmBIN.RS.EOF
    '全レコード処理が終わったらadFilterPendingRecordsでフィルタを掛けて結果があればUpdateBatchでDBに反映
    clsADOfrmBIN.RS.Filter = adFilterPendingRecords
    If clsADOfrmBIN.RS.RecordCount >= 1 Then
        '変更があればUpdateBatch
        clsADOfrmBIN.RS.UpdateBatch adAffectGroup
        If Not clsADOfrmBIN.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified Then
            '変更なしフラグが立っていなかったらメッセージ出す
            MsgBox "Statusの状態をDBに登録する際にエラーが発生しました"
            GoTo CloseAndExit
        End If
    End If
    clsADOfrmBIN.RS.Filter = adFilterNone
    '最初のレコードに戻りStatusCheckで色等を反映させて終わり
    clsADOfrmBIN.RS.MoveFirst
    'イベント再開する
    StopEvents = False
    'StatusCheck
    StatusCheck
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SetRsAllStatus code : " & Err.Number & " Description : " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBinLabel 
   Caption         =   "BINカードラベル印刷項目編集画面"
   ClientHeight    =   9105.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610.001
   OleObjectBlob   =   "frmBinLabel.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmBinLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'フォーム内共有変数
Private clsADOfrmBIN As clsADOHandle
Private clsEnumfrmBIN As clsEnum
Private clsSQLBc As clsSQLStringBuilder
Private dicObjNameToFieldName As Dictionary
Private clsIncrementalfrmBIN As clsIncrementalSerch
'メンバ変数
Private confrmBIN As ADODB.Connection
Private rsLabelTemp As ADODB.Recordset
Private rsOnlyPartsMaster As ADODB.Recordset                        '暫定対応になるかも？AddNew時に使用するPartsMasterのみをSelectした結果を格納(Addnewも)するRS
Private StopEvents As Boolean
Public UpdateMode As Boolean                                        '編集可能状態になってるときはTrueをセット
Private AddnewMode As Boolean                                       '新規追加モードの時にTrueをセット
Private strStartTime As String
Public strSavePoint As String                                       'LabelTempに追加する際のSavePoint文字列
Public varstrarrSelectedSavepoint As Variant                        '印刷対象となるSavePointを格納したstring配列
'定数
Private Const MAX_LABEL_TEXT_LENGTH As Long = 18
Private Const LABEL_TEMP_DELETE_FLAG As String = "LabelTempDelete"  'LabenTempテーブルを削除する時にStartTimeにセットする定数
'------------------------------------------------------------------------------------------------------
'SQL
'binLabelの基礎データ取得SQL
Private Const SQL_BIN_LABEL_DEFAULT_DATA As String = "SELECT TDBPrt.F_INV_Tehai_ID,TDBTana.F_INV_Tana_ID,TDBTana.F_INV_Tana_Local_Text as F_INV_Tana_Local_Text,TDBPrt.F_INV_Tehai_Code as F_INV_Tehai_Code,TDBPrt.F_INV_Label_Name_1 as F_INV_Label_Name_1,TDBPrt.F_INV_Label_Name_2 as F_INV_Label_Name_2,TDBPrt.F_INV_Label_Remark_1 as F_INV_Label_Remark_1,TDBPrt.F_INV_Label_Remark_2 as F_INV_Label_Remark_2,TDBTana.F_INV_Tana_System_Text as F_INV_Tana_System_Text," & vbCrLf & _
"TDBPrt.F_INV_Store_Code as F_INV_Store_Code " & vbCrLf & _
"FROM T_INV_M_Parts AS TDBPrt " & vbCrLf & _
"    INNER JOIN T_INV_M_Tana as TDBTana " & vbCrLf & _
"    ON TDBPrt.F_INV_Tana_ID = TDBTana.F_INV_Tana_ID"
'''binLabelにおいて、4701のデータのみに限定するSQL、他で既に使用している手配コードがあるため
Private Const SQL_BIN_LABEL_DEFAULT_DATA_ONLY_4701 As String = "SELECT TDBPrt.F_INV_Tehai_ID,TDBTana.F_INV_Tana_ID,TDBTana.F_INV_Tana_Local_Text as F_INV_Tana_Local_Text,TDBPrt.F_INV_Tehai_Code as F_INV_Tehai_Code,TDBPrt.F_INV_Label_Name_1 as F_INV_Label_Name_1,TDBPrt.F_INV_Label_Name_2 as F_INV_Label_Name_2,TDBPrt.F_INV_Label_Remark_1 as F_INV_Label_Remark_1,TDBPrt.F_INV_Label_Remark_2 as F_INV_Label_Remark_2,TDBTana.F_INV_Tana_System_Text as F_INV_Tana_System_Text," & vbCrLf & _
"TDBPrt.F_INV_Store_Code as F_INV_Store_Code " & vbCrLf & _
"FROM T_INV_M_Parts AS TDBPrt " & vbCrLf & _
"    INNER JOIN T_INV_M_Tana as TDBTana " & vbCrLf & _
"    ON TDBPrt.F_INV_Tana_ID = TDBTana.F_INV_Tana_ID " & vbCrLf & _
"    WHERE F_INV_Tana_System_Text LIKE ""BL%"" OR F_INV_Tana_System_Text LIKE ""K%"""
'新規追加時のSQL、ポイントはT_INV_N_TANAをRightJoinし、未登録の棚番もRSに含める点
'棚番リストはFilterでM_PartsでTana_IDがNullの物を抽出する
Private Const SQL_BIN_LABEL_ADDNEW_TEHAI_CODE As String = "SELECT TDBPrt.F_INV_Tehai_ID,TDBTana.F_INV_Tana_ID,TDBPrt.F_INV_Tana_ID AS TDBPrts_F_INV_Tana_ID,TDBTana.F_INV_Tana_Local_Text as F_INV_Tana_Local_Text,TDBPrt.F_INV_Tehai_Code as F_INV_Tehai_Code,TDBPrt.F_INV_Label_Name_1 as F_INV_Label_Name_1,TDBPrt.F_INV_Label_Name_2 as F_INV_Label_Name_2,TDBPrt.F_INV_Label_Remark_1 as F_INV_Label_Remark_1,TDBPrt.F_INV_Label_Remark_2 as F_INV_Label_Remark_2,TDBTana.F_INV_Tana_System_Text as F_INV_Tana_System_Text," & vbCrLf & _
"TDBPrt.F_INV_Store_Code as F_INV_Store_Code " & vbCrLf & _
"FROM T_INV_M_Parts AS TDBPrt " & vbCrLf & _
"    RIGHT JOIN T_INV_M_Tana as TDBTana " & vbCrLf & _
"    ON TDBPrt.F_INV_Tana_ID = TDBTana.F_INV_Tana_ID " & vbCrLf & _
"    WHERE TDBPrt.F_INV_Tana_ID IS NULL"
'AddNewでうまくいかなかったので、M_Parts単独のSelect文
Private Const SQL_BIN_LABEL_ONLY_PARTS As String = "SELECT TDBPrt.F_INV_Tehai_ID,TDBPrt.F_INV_Tana_ID,TDBPrt.F_INV_Tehai_Code as F_INV_Tehai_Code,TDBPrt.F_INV_Label_Name_1 as F_INV_Label_Name_1,TDBPrt.F_INV_Label_Name_2 as F_INV_Label_Name_2,TDBPrt.F_INV_Label_Remark_1 as F_INV_Label_Remark_1,TDBPrt.F_INV_Label_Remark_2 as F_INV_Label_Remark_2,InputDate," & vbCrLf & _
"TDBPrt.F_INV_Store_Code as F_INV_Store_Code " & vbCrLf & _
"FROM T_INV_M_Parts AS TDBPrt "
'印刷完了(SavePoint 選択済み)のデータを削除する
'{0}    INV_CONST.T_INV_LABEL_TEMP
'{1}    (MailMergeWhere)
Private Const SQL_DELET_COMP_DATA_FROM_LABEL_TEMP As String = "DELETE FROM {0} " & vbCrLf & _
"WHERE {1}"
Private Sub btnAddNewTehaiCode_Click()
    SwitchAddNewMode True
End Sub
'------------------------------------------------------------------------------------------------------
'イベント
'Form Initial
Private Sub UserForm_Initialize()
    'フォーム初期化時
    ConstRuctor
End Sub
'Form Terminate
Private Sub UserForm_Terminate()
    DestRuctor
End Sub
'Click
'編集制限解除
Private Sub btnEnableEdit_Click()
    SwitchtBoxEditmode True
End Sub
'最終的にDBにUpdateする
Private Sub btnDoUpdate_Click()
    DoUpdateBatch
End Sub
'変更を破棄
Private Sub btnCancelUpdate_Click()
    CancelUpdateBatch
End Sub
'新規識別名にチェンジ
Private Sub btnNewSavePoint_Click()
    Dim longMsgBoxReturn As Long
    longMsgBoxReturn = MsgBox("ラベル一時テーブル登録時の識別名を変更しますか？(出力先用紙を変更する時など)", vbYesNo)
    Select Case longMsgBoxReturn
    Case vbYes
        '変更する場合
        strSavePoint = ""
        SetSavePoint
        Exit Sub
    Case vbNo
        'キャンセルした場合
        MsgBox "キャンセルしました"
        Exit Sub
    End Select
End Sub
'ラベル一時テーブルに追加する
Private Sub btnAddnewLabelTemp_Click()
    Dim isCollect As Boolean
    isCollect = RecreateLabelTempTable
    If Not isCollect Then
        '一時テーブル作成に失敗
        MsgBox "一時テーブル作成に失敗したため、処理を中断します"
        Exit Sub
    End If
    'SavePointの設定
    SetSavePoint
    '次にカレントレコードをTempTableに追加する
    AddNewRStoLabelTemp
    'UpdateModeで編集状態破棄フラグ(チェックボックス)が立っていたら確認無しでキャンセルする(メモ書きで一時記入した場合等)
    If UpdateMode And chkBoxCancelUpdateModeatLabelTemp.Value Then
        'UpdateModeでなおかつ編集状態不可チェックボックスがTrueだった
        CancelUpdateBatch True
    End If
End Sub
'DBからデータを引っ張り、差し込み印刷の結果のDocを表示する
'ラベルプリンタ用BINカード表示ラベル
Private Sub btnCreateLabelDoc_Click()
    On Error GoTo ErrorCatch
    'clsadoを定義するが、DBPathを取得する位にしか使わないので、共有変数とは別に定義する
    Dim clsADOMailMerge As clsADOHandle
    Set clsADOMailMerge = CreateclsADOHandleInstance
    Dim fsoMailMerge  As FileSystemObject
    Set fsoMailMerge = New FileSystemObject
    'clsADOを明示的にデフォルトへ
    clsADOMailMerge.SetDBPathandFilenameDefault
    'ラベル印刷時はプリンタの設定が必要(イベント駆動)なので、Planeファイルも指定必須
    MailMergeDocCreate fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, INV_CONST.INV_DOC_LABEL_MAILMERGE), _
                    fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, INV_CONST.INV_DOC_LABEL_PLANE)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "btnCreateLabelDoc_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOMailMerge Is Nothing Then
        clsADOMailMerge.CloseClassConnection
        Set clsADOMailMerge = Nothing
    End If
    If Not fsoMailMerge Is Nothing Then
        Set fsoMailMerge = Nothing
    End If
    Exit Sub
End Sub
'現品票(小)作成
Private Sub btnCreateGenpinSmall_Click()
    On Error GoTo ErrorCatch
    'clsadoを定義するが、DBPathを取得する位にしか使わないので、共有変数とは別に定義する
    Dim clsADOMailMerge As clsADOHandle
    Set clsADOMailMerge = CreateclsADOHandleInstance
    Dim fsoMailMerge  As FileSystemObject
    Set fsoMailMerge = New FileSystemObject
    'clsADOを明示的にデフォルトへ
    clsADOMailMerge.SetDBPathandFilenameDefault
    'MailMerge実行
    MailMergeDocCreate fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, INV_CONST.INV_DOC_LABEL_GENPIN_SMALL)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "btnCreateGenpinSmall_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOMailMerge Is Nothing Then
        clsADOMailMerge.CloseClassConnection
        Set clsADOMailMerge = Nothing
    End If
    If Not fsoMailMerge Is Nothing Then
        Set fsoMailMerge = Nothing
    End If
    Exit Sub
End Sub
'スペックシート_詳細現品票(小)作成、表示
Private Sub btnCreateSpecSheetSmall_Click()
    On Error GoTo ErrorCatch
    'clsadoを定義するが、DBPathを取得する位にしか使わないので、共有変数とは別に定義する
    Dim clsADOMailMerge As clsADOHandle
    Set clsADOMailMerge = CreateclsADOHandleInstance
    Dim fsoMailMerge  As FileSystemObject
    Set fsoMailMerge = New FileSystemObject
    'clsADOを明示的にデフォルトへ
    clsADOMailMerge.SetDBPathandFilenameDefault
    'MailMerge実行
    MailMergeDocCreate fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, INV_CONST.INV_DOC_LABEL_SPECSHEET_Small)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "btnCreateSpecSheetSmall_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOMailMerge Is Nothing Then
        clsADOMailMerge.CloseClassConnection
        Set clsADOMailMerge = Nothing
    End If
    If Not fsoMailMerge Is Nothing Then
        Set fsoMailMerge = Nothing
    End If
    Exit Sub
End Sub
'''棚表示_矢印無し
Private Sub btnCreateTanaNoArrow_Click()
    On Error GoTo ErrorCatch
    'clsadoを定義するが、DBPathを取得する位にしか使わないので、共有変数とは別に定義する
    Dim clsADOMailMerge As clsADOHandle
    Set clsADOMailMerge = CreateclsADOHandleInstance
    Dim fsoMailMerge  As FileSystemObject
    Set fsoMailMerge = New FileSystemObject
    'clsADOを明示的にデフォルトへ
    clsADOMailMerge.SetDBPathandFilenameDefault
    'MailMerge実行
    MailMergeDocCreate fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, INV_CONST.INV_DOC_LABEL_TANA_NO_ARROW)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "btnCreateTanaNoArrow_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOMailMerge Is Nothing Then
        clsADOMailMerge.CloseClassConnection
        Set clsADOMailMerge = Nothing
    End If
    If Not fsoMailMerge Is Nothing Then
        Set fsoMailMerge = Nothing
    End If
    Exit Sub
End Sub
'手配コードをセットしたパーツマスター画面を表示する
Private Sub btnShowPMList_Click()
    If txtBox_F_INV_Tehai_Code.Text = "" Then
        Exit Sub
    End If
    Load frmINV_PartsMaster_List
    frmINV_PartsMaster_List.txtBox_F_INV_Tehai_Code.SetFocus
    frmINV_PartsMaster_List.txtBox_F_INV_Tehai_Code.Text = frmBinLabel.txtBox_F_INV_Tehai_Code.Text
    frmINV_PartsMaster_List.lstBox_Incremental.ListIndex = 0
    frmINV_PartsMaster_List.lstBox_Incremental.Visible = False
    frmINV_PartsMaster_List.Show
End Sub
'新規棚番登録画面表示
Private Sub btnRegistNewLocationfrmShow_Click()
    modCreateInstanceforAddin.ShowFrmRegistNewLocation
End Sub
'インクリメンタルリストClick
Private Sub lstBox_Incremental_Click()
    If (StopEvents Or UpdateMode) And Not AddnewMode Then
        'イベントストップかUpdateModeで、更にAddNewModeじゃない時は抜ける
        Exit Sub
    End If
    If AddnewMode And clsIncrementalfrmBIN.txtBoxRef.Name <> txtBox_F_INV_Tana_Local_Text.Name Then
        'AddnewModeの時、相手にするのは棚番テキストボックスの実なので、それ以外の場合は単純にリスト消去して終わり
        lstBox_Incremental.Visible = False
        Exit Sub
    End If
    If clsIncrementalfrmBIN.Incremental_LstBox_Click Then
        'この中に入ってる時点でRSにフィルターが適用されている
        'イベント停止
        StopEvents = True
        'AddnewModeの状態の応じて処理を分岐
        Select Case AddnewMode And clsIncrementalfrmBIN.txtBoxRef.Name = txtBox_F_INV_Tana_Local_Text.Name
        Case True
            'AddNewModeの時
            'AddNewの時はRSからデータ取得するのは棚番のみなのでここで直接値をセットしてしまう
            'Tana_Local
            txtBox_F_INV_Tana_Local_Text.Text = _
            clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(txtBox_F_INV_Tana_Local_Text.Name)).Value
            'Tana_System
            lbl_F_INV_Tana_System_Text.Caption = _
            clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(lbl_F_INV_Tana_System_Text.Name)).Value
            'ここでインクリメンタルリスト非表示にしてしまう
            lstBox_Incremental.Visible = False
        Case False
            '通常動作
            'RSから値取得
            GetValuFromRS
            '非表示はkeyupイベントで行うことにした
    '        'リストを非表示にする
    '        lstBox_Incremental.ListIndex = -1
    '        If lstBox_Incremental.ListCount >= 2 Then
    '            lstBox_Incremental.Visible = False
    '        Else
    '            lstBox_Incremental.Height = 0
    '        End If
            'ここまでで値の取得が完了しているので、通常の編集不可モードへ
            SwitchtBoxEditmode False
        End Select
        'イベント再開
        StopEvents = False
    End If
End Sub
'keyup
'インクリメンタルリストKeyUp
Private Sub lstBox_Incremental_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If AddnewMode Then
        '新規追加モードの時は抜ける
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    'インクリメンタルに丸投げ
    clsIncrementalfrmBIN.Incremental_LstBox_Key_UP KeyCode, Shift
    'インクリメンタルの結果、インクリメンタルリストが非表示になっていたらオーダーNoボックスにフォーカス
    If lstBox_Incremental.Visible = False Or lstBox_Incremental.Height = 0 Then
        'オーダーNoボックスにSetFocus
        txtBox_OrderNumber.SetFocus
    End If
    'イベント再開
    StopEvents = False
End Sub
'mouseup
'インクリメンタルリストMouseUP
Private Sub lstBox_Incremental_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '新規追加モードの時は抜ける
    If AddnewMode Then
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    'インクリメンタル
    clsIncrementalfrmBIN.Incremental_LstBox_Mouse_UP Button
    'インクリメンタルの結果、インクリメンタルリストが非表示になっていたらオーダーNoボックスにフォーカス
    If lstBox_Incremental.Visible = False Then
        'オーダーNoボックスにSetFocus
        txtBox_OrderNumber.SetFocus
    End If
    'イベント再開
    StopEvents = False
End Sub
'Change
'オーダーNoテキストボックス
Private Sub txtBox_OrderNumber_Change()
    If StopEvents Then
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'Ucase掛ける
    txtBox_OrderNumber.Text = UCase(txtBox_OrderNumber.Text)
    'イベント再開する
    StopEvents = False
End Sub
'手配コード
Private Sub txtBox_F_INV_Tehai_Code_Change()
    'イベント停止状態ではなく、更にアップデートモードでもないときにインクリメンタル実行
    If (StopEvents Or UpdateMode) And Not AddnewMode Then
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'テキストにUcaseかける
    If frmBinLabel.txtBox_F_INV_Tehai_Code.TextLength >= 1 Then
        frmBinLabel.txtBox_F_INV_Tehai_Code.Text = UCase(frmBinLabel.txtBox_F_INV_Tehai_Code.Text)
    End If
    Select Case AddnewMode
    Case True
        'AddNewModeの時(結果が0件になってもメッセージ表示せず、そのままリストを非表示にする)
        clsIncrementalfrmBIN.Incremental_TextBox_Change True, , True
        'RSに反映させる
        UpdateRSFromContrl ActiveControl
    Case False
        '通常モードの時(結果0件になったらメッセージ表示)
        'インクリメンタル実行
        clsIncrementalfrmBIN.Incremental_TextBox_Change False
    End Select
    'イベント再開する
    StopEvents = False
End Sub
'棚番
Private Sub txtBox_F_INV_Tana_Local_Text_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    'Ucase
    If frmBinLabel.txtBox_F_INV_Tana_Local_Text.TextLength >= 1 Then
        '文字が入力されていたらUcase掛ける
        'イベント停止
        StopEvents = True
        'Ucase
        frmBinLabel.txtBox_F_INV_Tana_Local_Text.Text = UCase(frmBinLabel.txtBox_F_INV_Tana_Local_Text.Text)
        'イベント再開
        StopEvents = False
    End If
    Select Case True
    Case (UpdateMode = True) And (AddnewMode = False)
        'アップデートモードで、なおかつAddNewModeじゃない時
        'RSに現在の値を設定する
        UpdateRSFromContrl ActiveControl
        Exit Sub
    Case (UpdateMode = False) Or (AddnewMode = True)
        '検索モード(?)の時か、AddNewModeの時(インクリメンタルを使用するモード)
        If TypeName(ActiveControl) <> "TextBox" Then
            'アクティブコントロールがテキストボックスじゃなかったら抜ける
            Exit Sub
        End If
        'イベント停止
        StopEvents = True
        'RSから内容を取得(listboxのClickイベントで呼ばれるはず)されるまでUpdateModeにしてはいけない
        'btnEnableEditをFalseに
        btnEnableEdit.Enabled = False
        'Ucase掛ける
        If ActiveControl.TextLength >= 1 Then
            ActiveControl.Text = UCase(ActiveControl.Text)
        End If
        Select Case AddnewMode
        Case True
            'AddNewModeの時、Changeしても他のテキストボックスの値を消去しない
            clsIncrementalfrmBIN.Incremental_TextBox_Change , , True
        Case False
            '通常動作
            clsIncrementalfrmBIN.Incremental_TextBox_Change
        End Select
        If AddnewMode And lstBox_Incremental.ListCount = 1 Then
            '新規追加モードで、なおかつ候補が残り1個になった場合
            '棚番ボックスにリストの値を設定し、リストを非表示にしてしまう
            txtBox_F_INV_Tana_Local_Text.Text = lstBox_Incremental.List(0)
            lstBox_Incremental.Visible = False
        End If
        'イベント再開
        StopEvents = False
        Exit Sub
    End Select
End Sub
'品名1
Private Sub txtBox_F_INV_Label_Name_1_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'RSのUpdateメソッドへ
    UpdateRSFromContrl ActiveControl
    'イベント再開する
    StopEvents = False
End Sub
'品名2
Private Sub txtBox_F_INV_Label_Name_2_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'RSのUpdateメソッドへ
    UpdateRSFromContrl ActiveControl
    'イベント再開する
    StopEvents = False
End Sub
'備考1
Private Sub txtBox_F_INV_Label_Remark_1_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'RSのUpdateメソッドへ
    UpdateRSFromContrl ActiveControl
    'イベント再開する
    StopEvents = False
End Sub
'備考2
Private Sub txtBox_F_INV_Label_Remark_2_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'RSのUpdateメソッドへ
    UpdateRSFromContrl ActiveControl
    'イベント再開する
    StopEvents = False
End Sub
'''貯蔵記号
Private Sub txtBox_F_INV_Store_Code_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'Ucase掛ける
    txtBox_F_INV_Store_Code.Text = UCase(txtBox_F_INV_Store_Code.Text)
    'RSのUpdateメソッドへ
    UpdateRSFromContrl ActiveControl
    'イベント再開する
    StopEvents = False
End Sub
'''4701限定チェックボックス
Private Sub chkBox_Only4701_Change()
    If StopEvents Or UpdateMode Or AddnewMode Then
        'イベント停止、UpdateMode、AddnewModeいずれかのフラグが立っていたら処理を中断
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    '全項目消去
    ClearAllContents
    'デフォルトデータ取得プロシージャ
    SetDefaultValuetoRS
    'イベント再開
    StopEvents = False
    '棚番ボックスにフォーカス(インクリメンタル初期化されるはず)
    txtBox_F_INV_Tana_Local_Text.SetFocus
    Exit Sub
End Sub
'Enter
'棚番テキストボックスEnter
Private Sub txtBox_F_INV_Tana_Local_Text_Enter()
    If (StopEvents Or UpdateMode) And Not AddnewMode Then
        'StopEvent か UpdateModeで、なおかつAddNewModeじゃなかったら抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    If AddnewMode And chkBoxShowUnUseLocationOnly.Value Then
        'AddNewModeでなおかつ未使用棚番のみ表示オプションがセットされていたら、最初に元データのSQLを変更し、データ再取得する
        If clsADOfrmBIN.RS.State = ObjectStateEnum.adStateOpen Then
            clsADOfrmBIN.RS.Close
        End If
        clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_ADDNEW_TEHAI_CODE
        clsADOfrmBIN.RS.Open
        clsADOfrmBIN.RS.Filter = ""
        If clsADOfrmBIN.RS.RecordCount < 1 Then
            '未使用の棚番が無いときはRecordCountが0になってるはず
            MsgBox "未使用の棚番がありませんでした。新規棚番登録ボタンから登録して下さい。"
            'AddNewMode解除
            SwitchAddNewMode False
            Exit Sub
        End If
        'インクリメンタル実行、Enter、フィルタ追加モード
        clsIncrementalfrmBIN.Incremental_TextBox_Enter txtBox_F_INV_Tana_Local_Text, lstBox_Incremental, True
    Else
        '通常動作
        'インクリメンタル実行、Enter
        clsIncrementalfrmBIN.Incremental_TextBox_Enter txtBox_F_INV_Tana_Local_Text, lstBox_Incremental, False
    End If
    'イベント再開
    StopEvents = False
    Exit Sub
End Sub
Private Sub txtBox_F_INV_Tehai_Code_Enter()
    'イベント停止状態ではなく、更にアップデートモードでもないときにインクリメンタル実行
    If (StopEvents Or UpdateMode) And Not AddnewMode Then
        'StopEventsかUpdateModeでなおかつAddnewModeじゃない時は抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'インクリメンタル実行、リストを表示するのが目的
    If AddnewMode Then
        'AddNewModeの時はSQLを一旦通常のものに変えてRequeryする
        If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
            clsADOfrmBIN.RS.Close
        End If
        clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_DEFAULT_DATA
        clsADOfrmBIN.RS.Open
    End If
    clsIncrementalfrmBIN.Incremental_TextBox_Enter frmBinLabel.txtBox_F_INV_Tehai_Code, frmBinLabel.lstBox_Incremental
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'インクリメンタルリストEnter
Private Sub lstBox_Incremental_Enter()
    If StopEvents Or UpdateMode Then
        'StopEbentかUpdateModeの時は抜ける
        Exit Sub
    End If
    '残り候補が1個だったらClickイベント発生させるだけ
    clsIncrementalfrmBIN.Incremental_LstBox_Enter
    Exit Sub
End Sub
'------------------------------------------------------------------------------------------------------
'メソッド
'''コンストラクタ
Private Sub ConstRuctor()
    'インスタンス共有変数の初期化
    If clsADOfrmBIN Is Nothing Then
        Set clsADOfrmBIN = CreateclsADOHandleInstance
    End If
    If clsEnumfrmBIN Is Nothing Then
        Set clsEnumfrmBIN = CreateclsEnum
    End If
    If clsSQLBc Is Nothing Then
        Set clsSQLBc = CreateclsSQLStringBuilder
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
    'とりあえずイベントは停止状態にする
    StopEvents = True
    '初回データ設定
    SetDefaultValuetoRS
    'objToFieldNameを設定
    setObjToFieldNameDic
    clsIncrementalfrmBIN.ConstRuctor Me, dicObjNameToFieldName, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc
    'RSのデータを取得する
    'ここでは取得しないで、インクリメンタルサーチに任せる
'    GetValuFromRS
    'ラベル追加枚数コンボボックスの初期化
    Dim longRowCounter As Long
    Dim longarrAddCount(9) As Long
    For longRowCounter = 0 To 9
        longarrAddCount(longRowCounter) = longRowCounter + 1
    Next longRowCounter
    'オーダーNoテキストボックスのMaxLength設定する目的でLabelTempテーブル作成プロシージャを走らせる
    RecreateLabelTempTable
    cmbBox_AddLabelCount.List = longarrAddCount
    'イベント再開する
    StopEvents = False
#If DebugDB Then
    MsgBox "DebugDB有効"
#End If
End Sub
'''デストラクタ
Private Sub DestRuctor()
    'メンバ変数の解放、特に接続が関連しているものは重点的に
    If Not clsADOfrmBIN.RS Is Nothing Then
        clsADOfrmBIN.RS.ActiveConnection.Close
'        clsADOfrmBIN.RS.Close
        Set clsADOfrmBIN.RS = Nothing
    End If
    If Not clsADOfrmBIN Is Nothing Then
        clsADOfrmBIN.CloseClassConnection
        Set clsADOfrmBIN = Nothing
    End If
    If Not confrmBIN Is Nothing Then
        If confrmBIN.State And ObjectStateEnum.adStateOpen Then
            '接続していたら閉じる
            confrmBIN.Close
        End If
        Set confrmBIN = Nothing
    End If
    If Not clsIncrementalfrmBIN Is Nothing Then
        Set clsIncrementalfrmBIN = Nothing
    End If
    Me.Hide
    Unload Me
End Sub
'''メンバ変数のRecordSetに初期データを設定する
'''args
'''AddNewMode(クラス変数)      Trueがセットされていたら、途中のJoinがRight Joinになり、未使用の棚番もRSに含まれるようになる
Private Sub SetDefaultValuetoRS()
    '最初にclsadoのDBPathとDBFilnameをデフォルトに
    clsADOfrmBIN.SetDBPathandFilenameDefault
    'もし接続されていたら切断する
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        clsADOfrmBIN.RS.Close
    End If
    If confrmBIN.State And ObjectStateEnum.adStateOpen Then
        confrmBIN.Close
    End If
    'Connectionの設定をする
    confrmBIN.ConnectionString = clsADOfrmBIN.CreateConnectionString(clsADOfrmBIN.DBPath, clsADOfrmBIN.DBFileName)
    confrmBIN.CursorLocation = adUseClient
    confrmBIN.Mode = adModeRead Or adModeShareDenyNone
    '接続オープン
    confrmBIN.Open
    'RSのプロパティを設定していく
    clsADOfrmBIN.RS.LockType = adLockBatchOptimistic
    clsADOfrmBIN.RS.CursorType = adOpenStatic
    'rsのSourceにSQL設定(後でパラメータ対応する)
    'AddNewModeにより処理を分岐
    Select Case AddnewMode
    Case True
        '新規追加モード
        'Tana_がRightJoinになる
        clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_ADDNEW_TEHAI_CODE
    Case False
        '通常動作
        '4701限定フラグの有無でSQLを分ける
        If chkBox_Only4701.Value Then
            '4701限定フラグが立っていた場合
            clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_DEFAULT_DATA_ONLY_4701
        Else
            '全てのデータを対象にする場合
            clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_DEFAULT_DATA
        End If
    End Select
    'rsのActiveConnectionにConnectionオブジェクト指定
    Set clsADOfrmBIN.RS.ActiveConnection = confrmBIN
    'rsオープン
    clsADOfrmBIN.RS.Open , , , , CommandTypeEnum.adCmdText
    '以下は正常に動く
    '更新に必要なキー列の情報が～・・・→両方のテーブルの主キーをSELECTのフィールドに含めると解決
'    clsADOfrmBIN.RS.Fields("F_INV_Label_Name_2").Value = "InputTest"
'    clsADOfrmBIN.RS.Fields("F_INV_Tana_Local_Text").Value = "K23 A01"
'    clsADOfrmBIN.RS.Update
'    clsADOfrmBIN.RS.UpdateBatch
    DebugMsgWithTime "Default Data count: " & clsADOfrmBIN.RS.RecordCount
End Sub
'dicobjToFieldNameの設定
Private Sub setObjToFieldNameDic()
    If dicObjNameToFieldName Is Nothing Then
        Set dicObjNameToFieldName = New Dictionary
    End If
    '最初に全消去
    dicObjNameToFieldName.RemoveAll
    '項目を追加していく
    '今回はテーブル毎にフィールド名が独立しているので、テーブルプリフィックスは無しでRSで格納している
    dicObjNameToFieldName.Add txtBox_F_INV_Tana_Local_Text.Name, clsEnumfrmBIN.INVMasterTana(F_INV_Tana_Local_Text_IMT)
    dicObjNameToFieldName.Add txtBox_F_INV_Tehai_Code.Name, clsEnumfrmBIN.INVMasterParts(F_Tehai_Code_IMPrt)
    dicObjNameToFieldName.Add txtBox_F_INV_Label_Name_1.Name, clsEnumfrmBIN.INVMasterParts(F_Label_Name_1_IMPrt)
    dicObjNameToFieldName.Add txtBox_F_INV_Label_Name_2.Name, clsEnumfrmBIN.INVMasterParts(F_Label_Name_2_IMPrt)
    dicObjNameToFieldName.Add txtBox_F_INV_Label_Remark_1.Name, clsEnumfrmBIN.INVMasterParts(F_Label_Remark_1_IMPrt)
    dicObjNameToFieldName.Add txtBox_F_INV_Label_Remark_2.Name, clsEnumfrmBIN.INVMasterParts(F_Label_Remark_2_IMPrt)
    dicObjNameToFieldName.Add lbl_F_INV_Tana_System_Text.Name, clsEnumfrmBIN.INVMasterTana(F_INV_Tana_System_Text_IMT)
    dicObjNameToFieldName.Add txtBox_F_INV_Store_Code.Name, clsEnumfrmBIN.INVMasterParts(F_Store_Code_IMPrt)
End Sub
'cidObjToFieldにあるコントロールの値をすべて消去する
Private Sub ClearAllContents()
    'イベント停止する
    StopEvents = True
    Dim varKeyobjDic As Variant
    'dicObjtoFieldループ
    For Each varKeyobjDic In dicObjNameToFieldName
        Select Case TypeName(Me.Controls(varKeyobjDic))
        Case "TextBox"
            'TextBoxだった時
            Me.Controls(varKeyobjDic).Text = ""
        Case "Label"
            'Labelだった時
            Me.Controls(varKeyobjDic).Caption = ""
        End Select
    Next varKeyobjDic
    '単独でオーダーNo消去
    txtBox_OrderNumber.Text = ""
    '枚数リストボックスListIndexを0(1枚)に
    cmbBox_AddLabelCount.ListIndex = 0
End Sub
'RSから値をとってくる
Private Sub GetValuFromRS()
    On Error GoTo ErrorCatch
    If UpdateMode Then
        'UpdateModだったら抜ける
        GoTo CloseAndExit
    End If
    If clsADOfrmBIN.RS.EOF And clsADOfrmBIN.RS.BOF Then
            'BOFとEOF両方同時に立っていたらレコードが無いので抜ける
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    '一旦全項目消去
    ClearAllContents
    Dim varKeyobjDic As Variant
    'dicObjtoFieldをループ
    For Each varKeyobjDic In dicObjNameToFieldName
        Select Case True
        Case IsNull(clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(varKeyobjDic)).Value)
            'Nullだった場合
            'とりあえず空文字にする
            Select Case TypeName(Me.Controls(varKeyobjDic))
            Case "TextBox"
                'テキストボックスだったら
                Me.Controls(varKeyobjDic).Text = ""
            Case "Label"
                'ラベルだった
                Me.Controls(varKeyobjDic).Caption = ""
            End Select
        Case Else
            'データがあった場合
            'RSのデータをそのまま適用する
            Select Case TypeName(Me.Controls(varKeyobjDic))
            Case "TextBox"
                'テキストボックス
                Me.Controls(varKeyobjDic).Text = clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(varKeyobjDic)).Value
            Case "Label"
                'ラベル
                Me.Controls(varKeyobjDic).Caption = clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(varKeyobjDic)).Value
            End Select
        End Select
    Next varKeyobjDic
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "GetValuFromRS code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開
    StopEvents = False
    Exit Sub
End Sub
'''レコードを進んだり戻ったりする
'''args
'''intargKeyCode    基本はキー操作にする、→で次へ、←で前へ
Private Sub MoveRecord(intargKeyCode As Integer)
    If clsADOfrmBIN.RS.BOF And clsADOfrmBIN.RS.EOF Then
        'BOFとEOF両方立ってたら抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    Select Case intargKeyCode
    Case vbKeyRight
        '右、次へ
        clsADOfrmBIN.RS.MoveNext
        If clsADOfrmBIN.RS.EOF Then
            MsgBox "現在のレコードが最終レコードです"
            clsADOfrmBIN.RS.MovePrevious
        End If
    Case vbKeyLeft
        '左、前へ
        clsADOfrmBIN.RS.MovePrevious
        If clsADOfrmBIN.RS.BOF Then
            MsgBox "現在のレコードが先頭レコードです"
            clsADOfrmBIN.RS.MoveNext
        End If
    End Select
    '値の取得をする
    GetValuFromRS
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "MoveRecord code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'''テキストボックスの編集可能状態を切り替える
'''args
'''Editable     Trueにセットすると変更可能に、Falseで変更不可にする
Private Sub SwitchtBoxEditmode(Editable As Boolean)
    Select Case Editable
    Case True
        '編集可能にするとき
        UpdateMode = True
        btnDoUpdate.Enabled = True
        btnCancelUpdate.Enabled = True
        chkBox_Only4701.Enabled = False
        '手配コードテキストボックスはLockedにする
        txtBox_F_INV_Tehai_Code.Locked = True
        'LockedをFalseにして、BackColoreを薄緑にする
        txtBox_F_INV_Tana_Local_Text.Locked = False
        txtBox_F_INV_Tana_Local_Text.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Name_1.Locked = False
        txtBox_F_INV_Label_Name_1.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Name_2.Locked = False
        txtBox_F_INV_Label_Name_2.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Remark_1.Locked = False
        txtBox_F_INV_Label_Remark_1.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Remark_2.Locked = False
        txtBox_F_INV_Label_Remark_2.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Store_Code.Locked = False
        txtBox_F_INV_Store_Code.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        '編集可能設定ボタンを無効に
        btnEnableEdit.Enabled = False
    Case False
        '編集不可にするとき
        UpdateMode = False
        chkBox_Only4701.Enabled = True
        'UpdateBatckボタンをFalseに
        btnDoUpdate.Enabled = False
        btnCancelUpdate.Enabled = False
        '手配コードテキストボックスのLockedを解除する(インクリメンタル向けに入力できるようにする)
        txtBox_F_INV_Tehai_Code.Locked = False
        'LockedをTrueにして、BackColoreを標準背景色にする
        '棚番テキストボックスは編集不可モードの時はインクリメンタルに使うのでLockはしない
'        txtBox_F_INV_Tana_Local_Text.Locked = True
        txtBox_F_INV_Tana_Local_Text.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Name_1.Locked = True
        txtBox_F_INV_Label_Name_1.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Name_2.Locked = True
        txtBox_F_INV_Label_Name_2.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Remark_1.Locked = True
        txtBox_F_INV_Label_Remark_1.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Remark_2.Locked = True
        txtBox_F_INV_Label_Remark_2.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Store_Code.Locked = True
        txtBox_F_INV_Store_Code.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        '編集可能設定ボタンを有効に
        btnEnableEdit.Enabled = True
    End Select
End Sub
'''新規追加モードと通常モードをスイッチする
'''このメソッドを呼ぶたびにRSのデータはリセットされる
'''args
'''IsAddNewMode     Trueにセットすると新規追加モード、Falseにすると通常モード
Private Sub SwitchAddNewMode(IsAddNewMode As Boolean)
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    'clsIncrementalのイベントも一時停止する
    clsIncrementalfrmBIN.StopEvent = True
    'インクリメンタルリストのVisibleもFalseに
    lstBox_Incremental.Visible = False
    Select Case IsAddNewMode
    Case True
        '新規追加モードにする場合
        'AddNewフラグを立てる
        AddnewMode = True
        'UpdateModeをセットする
        SwitchtBoxEditmode True
        '追加で手配コードボックスも編集可能にする
        txtBox_F_INV_Tehai_Code.Locked = False
        txtBox_F_INV_Tehai_Code.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
        '新規追加モードボタンEnabledをFalseに
        btnAddNewTehaiCode.Enabled = False
        '未使用棚番チェックボックスEnabled True
        chkBoxShowUnUseLocationOnly.Enabled = True
        '4701限定チェックボックスEnable
        chkBox_Only4701.Enabled = False
        '全項目消去
        ClearAllContents
        'DBよりデータ再取得
        SetDefaultValuetoRS
        '一旦フォーカスを棚番テキストボックスから外す
        txtBox_F_INV_Tehai_Code.SetFocus
    Case False
        '通常モードにする場合
        'AddNewフラグを下げる
        AddnewMode = False
        SwitchtBoxEditmode False
        '新規追加モードボタンを使用可能に
        btnAddNewTehaiCode.Enabled = True
        '未使用棚番チェックボックスEnabled False
        chkBoxShowUnUseLocationOnly.Enabled = False
        '4701限定チェックボックスEnable
        chkBox_Only4701.Enabled = True
        '手配コードボックスを編集不可に戻す
        'インクリメンタルで使用するのでLockedはそのまま
        '色だけ戻す
        txtBox_F_INV_Tehai_Code.BackColor = FormCommon.TXTBOX_BACKCOLORE_NORMAL
        '全項目消去
        ClearAllContents
        'DBよりデータ再取得
        SetDefaultValuetoRS
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SwitchAddNewMode code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    '棚番テキストボックスにフォーカスセット(イニシャル値セットされるはず)
    txtBox_F_INV_Tana_Local_Text.SetFocus
    'clsIncrementalのイベントも再開する
    clsIncrementalfrmBIN.StopEvent = False
    Exit Sub
End Sub
''''各コントロールの値をRSにセットする
'''args
'''Optional rsargOnlyParts          AddNewModeの時は必須、PartsMasterオンリーのSelect文に対応するRS、通常のUpdateModeでは使用しない・・・はず
'''Optional rsargOnlyTana           オプション。そのうち棚番も一緒に新規登録するようになったら棚番オンリーのRSが必要になるかも?予約枠
Private Sub UpdateRSFromContrl(argCtrl As Control, Optional ByRef rsargOnlyParts As ADODB.Recordset, Optional ByRef rsargOnlyTana As ADODB.Recordset)
    On Error GoTo ErrorCatch
    If AddnewMode Then
        'AddNewModeの時,rsOnlyPartsとrsOnlyTanaの初期化状態をチェック、新規レコードチェックフラグON
        rsOnlyPartsInitialize True
        Set rsargOnlyParts = rsOnlyPartsMaster
    End If
    If Not dicObjNameToFieldName.Exists(argCtrl.Name) Then
        'dicobjToFieldに存在しないコントロール名の場合は抜ける
        Exit Sub
    End If
    If AddnewMode And (rsargOnlyParts Is Nothing) And (rsargOnlyTana Is Nothing) Then
        '新規追加モードで、個別RSがどちらもNothingだったら抜ける
        MsgBox "RecordSetが未初期化でした。処理を中断します"
        Exit Sub
    End If
    '参照用RS定義
    Dim rsRefLocal As ADODB.Recordset
    '以下、各条件に応じてRSの参照元を決定
    Select Case True
    Case UpdateMode And Not AddnewMode
        'UpdateModeの時
        '対象にするRSはclsADOのもの
        Set rsRefLocal = clsADOfrmBIN.RS
    Case AddnewMode
        'AddNewMode
        Select Case True
        Case clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).Properties("BASETABLENAME") = INV_CONST.T_INV_M_Parts
            'AddnewModeでベーステーブル名がPatrsだった
            Set rsRefLocal = rsargOnlyParts
        Case clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).Properties("BASETABLENAME") = INV_CONST.T_INV_M_Tana
            'AddnewModeでベーステーブルがTanaだった
            Set rsRefLocal = rsargOnlyTana
        End Select
    End Select          'ModeSelector
    If rsRefLocal Is Nothing Then
        DebugMsgWithTime "UpdateRSRromContrl : refRS is nothing"
        GoTo CloseAndExit
    End If
    '取得したRS参照に対して処理を行う
    Select Case True
    '最初に文字数チェックを行い、オーバーしていたら設定値まで切り下げる
    Case Len(argCtrl.Text) > rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize
        '文字数がフィールド設定値オーバー
        MsgBox "入力された文字数が設定の " & rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize & " 文字を超えています。"
        argCtrl.Text = Mid(argCtrl.Text, 1, rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize)
        '切り下げたデータをRSに登録する
        rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).Value = _
        argCtrl.Text
        rsRefLocal.Update
    Case IsNull(rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).Value), rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).Value <> argCtrl.Text
        'RSの値がNullか、引数のコントロールのtextと違っている場合
        'rsに値をセットして、Updateまでする（DBに反映するにはUpdateBatchしないとダメ）
        rsRefLocal.Fields(dicObjNameToFieldName(argCtrl.Name)).Value = _
        argCtrl.Text
        rsRefLocal.Update
    End Select          'CheckDigit
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "UpdateRSFromContrl code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'UpdateでRSにコミットされた変更をDBにプッシュする
Private Sub DoUpdateBatch()
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    If AddnewMode Then
        'AddNewModeの時は別プロシージャへ(帰ってこない)
        AddnewUpdateDB
        Exit Sub
    End If
    Dim varOldFilter As Variant
    Dim varBookMark As Variant
    '現在のBookMarkを取得
    varBookMark = clsADOfrmBIN.RS.Bookmark
    '古いフィルターを退避
    varOldFilter = clsADOfrmBIN.RS.Filter
    '一旦フィルタ解除する
    clsADOfrmBIN.RS.Filter = adFilterNone
    'rsのFilterを adFilterPendingRecords、サーバーに未送信の変更のあるレコードだけ、にする
    clsADOfrmBIN.RS.Filter = adFilterPendingRecords
    Select Case True
    Case (clsADOfrmBIN.RS.BOF And clsADOfrmBIN.RS.EOF), clsADOfrmBIN.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified
        'adFilterPendingRecords掛けた後レコードがなかったか、Unmodifiedフラグが立っていた場合
        MsgBox "変更点はありませんでした。"
        'フィルタを戻してやる
        clsADOfrmBIN.RS.Filter = varOldFilter
        '通常モード(編集不可モード)に戻す
        SwitchtBoxEditmode False
        GoTo CloseAndExit
    Case clsADOfrmBIN.RS.Status And adRecModified
        'UpdateBatckで引数を与えないとうまく更新できないことがあるみたいなので、adAffectGroup、rs.filter(定数)で抽出されたレコードだけに影響あるやつで
        'adAffectCurrentでFilterで指定されたレコードのみ更新する引数を指定
        clsADOfrmBIN.RS.UpdateBatch adAffectGroup
'        'Filterを戻してやる
'        clsADOfrmBIN.RS.Filter = varOldFilter
        'BookMarkを戻す
        clsADOfrmBIN.RS.Bookmark = varBookMark
        If (clsADOfrmBIN.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified) Then
            MsgBox "正常に更新されました"
            'FilterをNoneにする
            clsADOfrmBIN.RS.Filter = adFilterNone
            '編集不可モードへ
            SwitchtBoxEditmode False
'            'RSよりデータを取得する
'            GetValuFromRS
            'BookMarkを戻す
            clsADOfrmBIN.RS.Bookmark = varBookMark
            GoTo CloseAndExit
        Else
            MsgBox "更新に失敗した可能性があります RSStasus: " & clsADOfrmBIN.RS.Status
            GoTo CloseAndExit
        End If
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "DoUpdateBatch code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'''AddNewModeでDBをUpdateする
Private Sub AddnewUpdateDB()
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    '手配コードと棚番の組み合わせで同一のものがないかチェックする
    'SQLを一旦標準のものへ
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        '接続されていたら一旦切断する
        clsADOfrmBIN.RS.Close
    End If
    clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_DEFAULT_DATA
    clsADOfrmBIN.RS.Open , , , , CommandTypeEnum.adCmdText
    Select Case True
    End Select
    'RSの初期化、AddNewStatusの確認
    Select Case True
    Case rsOnlyPartsMaster Is Nothing, Not CBool(rsOnlyPartsMaster.Status And ADODB.RecordStatusEnum.adRecNew)
        'rsOnlyが未初期化か、NewRecordのフラグが立っていなかったら抜ける
        MsgBox "登録用rsが未初期化、または新規レコードが見つかりません"
        GoTo CloseAndExit
    End Select
    'Filterをセットしてやる
    clsADOfrmBIN.RS.Filter = dicObjNameToFieldName(txtBox_F_INV_Tana_Local_Text.Name) & " = '" & txtBox_F_INV_Tana_Local_Text.Text & "' AND " & _
                            dicObjNameToFieldName(txtBox_F_INV_Tehai_Code.Name) & " = '" & txtBox_F_INV_Tehai_Code.Text & "'"
    If clsADOfrmBIN.RS.RecordCount >= 1 Then
        DebugMsgWithTime "AddnewUpdateDB : Already exist Tehaicode and LocalTana pair"
        MsgBox "指定の手配コードと棚番の組み合わせは既に存在します" & vbCrLf & _
            "手配コード: " & txtBox_F_INV_Tehai_Code.Text & vbCrLf & _
            "棚番： " & txtBox_F_INV_Tana_Local_Text.Text
        GoTo CloseAndExit
    End If
    'SQLをAddNewModeへ
    If clsADOfrmBIN.RS.State And ObjectStateEnum.adStateOpen Then
        clsADOfrmBIN.RS.Close
    End If
    clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_ADDNEW_TEHAI_CODE
    clsADOfrmBIN.RS.Open , , , , CommandTypeEnum.adCmdText
    'フィルターに棚番のテキストを設定
    clsADOfrmBIN.RS.Filter = dicObjNameToFieldName(txtBox_F_INV_Tana_Local_Text.Name) & " = '" & txtBox_F_INV_Tana_Local_Text.Text & "'"
    If clsADOfrmBIN.RecordCount < 1 Then
        MsgBox "指定の棚番が見つかりませんでした。棚番新規登録画面で登録しなおしてみて下さい。"
        GoTo CloseAndExit
    ElseIf clsADOfrmBIN.RecordCount > 1 Then
        MsgBox "指定の棚番で複数のレコードがありました。DBのメンテナンスが必要です。"
        GoTo CloseAndExit
    End If
    '登録作業に入る
    'tana_IDを退避
    Dim longTanaID As Long
    longTanaID = clsADOfrmBIN.RS.Fields(clsEnumfrmBIN.INVMasterTana(F_INV_TANA_ID_IMT)).Value
    'rsOnlyPartsの初期化確認
    rsOnlyPartsInitialize
    'PartsOnlyに退避したTanaIDを設定
    rsOnlyPartsMaster.Fields(clsEnumfrmBIN.INVMasterParts(F_Tana_ID_IMPrt)).Value = longTanaID
    'InputDate入力
    rsOnlyPartsMaster.Fields(PublicConst.INPUT_DATE).Value = GetLocalTimeWithMilliSec
    'RSを確定
    rsOnlyPartsMaster.Update
    'RSのフィルタを再設定、定数のものへ
    rsOnlyPartsMaster.Filter = adFilterNone
    rsOnlyPartsMaster.Filter = adFilterPendingRecords
    If Not CBool(rsOnlyPartsMaster.BOF And rsOnlyPartsMaster.EOF) And CBool(rsOnlyPartsMaster.Status And ADODB.RecordStatusEnum.adRecNew) Then
        'レコードが存在し、なおかつRSの状態が新規レコード有の場合
        rsOnlyPartsMaster.UpdateBatch adAffectGroup
    End If
    If rsOnlyPartsMaster.Status And ADODB.RecordStatusEnum.adRecUnmodified Then
        MsgBox "正常に追加されました"
        '改めてイベント停止する
        StopEvents = True
        '通常モードへ戻す
        SwitchAddNewMode False
        'イベント再開する
        StopEvents = False
    Else
        MsgBox "追加に失敗しました"
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "AddnewUpdateDB code: " & Err.Number & " Description: " & Err.Description
    MsgBox "登録時にエラーが発生しました " & vbCrLf & Err.Description
    rsOnlyPartsMaster.CancelUpdate
    rsOnlyPartsMaster.CancelBatch
    GoTo CloseAndExit
CloseAndExit:
    '一時的に使用したRSonlyの接続を切断する
    If rsOnlyPartsMaster.State And ObjectStateEnum.adStateOpen Then
        rsOnlyPartsMaster.Close
        Set rsOnlyPartsMaster = Nothing
    End If
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'''rsOnlyPartsの初期化を行う
'''args
'''Optional CheckNewRecord      Trueをセットすると新規レコード追加されていない時はAddNewまでやってしまう
Private Sub rsOnlyPartsInitialize(Optional CheckNewRecord As Boolean)
    On Error GoTo ErrorCatch
    If rsOnlyPartsMaster Is Nothing Then
        '初期化されてなかったら
        Set rsOnlyPartsMaster = New ADODB.Recordset
    End If
    If Not CBool(rsOnlyPartsMaster.State And ObjectStateEnum.adStateOpen) Then
        '接続されていなかったら初回接続を行う
        '登録用にPartsMasterのみのSelectでRecordSetを新たに設定
        'Connectionはクラス共有変数の物を設定
        Set rsOnlyPartsMaster.ActiveConnection = clsADOfrmBIN.RS.ActiveConnection
        'PartsMasterオンリーのSQLをセット
        rsOnlyPartsMaster.Source = SQL_BIN_LABEL_ONLY_PARTS
        rsOnlyPartsMaster.LockType = adLockBatchOptimistic
        rsOnlyPartsMaster.CursorType = adOpenStatic
        rsOnlyPartsMaster.CursorLocation = adUseClient
        'RSオープン
        rsOnlyPartsMaster.Open , , , , CommandTypeEnum.adCmdText
    End If
    If CheckNewRecord Then
        '新規レコードチェックフラグが立っていたら
        Select Case True
        Case Not CBool(rsOnlyPartsMaster.Status And ADODB.RecordStatusEnum.adRecNew)
            'rsに新規レコードフラグが立っていなかった場合
            'AddNewする
            rsOnlyPartsMaster.AddNew
        End Select
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "rsOnlyPartsInitialize code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
''''変更された内容を破棄して元に戻す
'''args
'''Optional NoConfirm           Trueをセットすると確認無しでキャンセルする、デフォルトはFalse
Private Sub CancelUpdateBatch(Optional NoConfirm As Boolean = False)
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    '現在のBookMarkを退避
    If clsADOfrmBIN.RS.Supports(adBookmark) Then
        'BookMarkが有効だったら
        Dim varBookMark As Variant
        varBookMark = clsADOfrmBIN.RS.Bookmark
    End If
    '今のフィルタを退避
    Dim varOldFilter As Variant
    varOldFilter = clsADOfrmBIN.RS.Filter
    '一旦フィルタ解除
    clsADOfrmBIN.RS.Filter = adFilterNone
    '新しく adfilterPendingRecordsで変更点のあるレコードだけに絞り込む
    clsADOfrmBIN.RS.Filter = adFilterPendingRecords
    Select Case True
    Case clsADOfrmBIN.RS.BOF And clsADOfrmBIN.RS.EOF, clsADOfrmBIN.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified
        'EOR,BOFが同時に立ってるか(変更レコード無し)、StatusがUnmodifiedになっているとき
        If Not NoConfirm Then
            MsgBox "変更点はありませんでした"
        End If
        'フィルタを戻す？
        clsADOfrmBIN.RS.Filter = varOldFilter
        If clsADOfrmBIN.RS.Supports(adBookmark) Then
            'BookMark有効ならBookMarkを戻す
            clsADOfrmBIN.RS.Bookmark = varBookMark
        End If
        '編集不可モードへ
        SwitchtBoxEditmode False
        If AddnewMode Then
            'AddNewModeの時はこっちも解除してやる
            SwitchAddNewMode False
        End If
        GoTo CloseAndExit
    Case clsADOfrmBIN.RS.Status And adRecModified
        'Statusが変更有になっている
        'キャンセルしていいか問い合わせ
        Dim longMsgBoxRet As Long
        If Not NoConfirm Then
            longMsgBoxRet = MsgBox("内容が変更されています、変更を破棄しても良いですか?", vbYesNo)
        End If
        If longMsgBoxRet = vbNo Then
            'キャンセルされた
            MsgBox "変更の破棄をキャンセルしました。データは変更後のままです。"
            'フィルタを戻す
            clsADOfrmBIN.RS.Filter = varOldFilter
            If clsADOfrmBIN.RS.Supports(adBookmark) Then
                'BookMark有効ならBookMarkを戻す
                clsADOfrmBIN.RS.Bookmark = varBookMark
            End If
            Exit Sub
        End If
        'フィルタ後のレコード限定でCancelBatch
        clsADOfrmBIN.RS.CancelBatch adAffectGroup
'        If (clsADOfrmBIN.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified) Or (clsADOfrmBIN.RS.Status = ADODB.RecordStatusEnum.adRecOK) Then
        If (clsADOfrmBIN.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified) Then
            If Not NoConfirm Then
                MsgBox "変更点を無事に破棄しました。"
            End If
            'フィルタを戻す？
            clsADOfrmBIN.RS.Filter = varOldFilter
            If clsADOfrmBIN.RS.Supports(adBookmark) Then
                'BookMark有効ならBookMarkを戻す
                clsADOfrmBIN.RS.Bookmark = varBookMark
            End If
            '編集不可モードへ
            SwitchtBoxEditmode False
            If AddnewMode Then
                'AddNewModeの時はこっちも解除してやる
                SwitchAddNewMode False
            End If
            'RSより値を取得する
            GetValuFromRS
            GoTo CloseAndExit
        Else
            MsgBox "変更の破棄に失敗した可能性があります RSStasus: " & clsADOfrmBIN.RS.Status
            GoTo CloseAndExit
        End If
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "CancelUpdateBatch code: " & Err.Number & " Description: " & Err.Description
    If Err.Number = -2147217906 Then
        'ブックマークが無効ですのエラーの時は(変更後等でブックマークが移動)編集不可モードへ戻してやる
        SwitchtBoxEditmode False
    End If
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'''ラベル出力用一時テーブルを作成する
'''既存のテーブルが存在していたら強制的に削除してから新たに作成する
Private Function RecreateLabelTempTable() As Boolean
    On Error GoTo ErrorCatch
    '以下の操作は独立してConnection張りたいので、クラス共有clsADOインスタンスは使用しない
    Dim clsADOLabelTemp As clsADOHandle
    Set clsADOLabelTemp = CreateclsADOHandleInstance
    'DBPathはデフォルト、DBFilenameは一時テーブル格納DBのものにする
    clsADOLabelTemp.SetDBPathandFilenameDefault
    clsADOLabelTemp.DBFileName = PublicConst.TEMP_DB_FILENAME
'    'SavePoint導入により、テーブル削除は別プロシージャで判断する
    'テーブル存在チェックし、無かったら作成する
    If Not clsADOLabelTemp.IsTableExists(INV_CONST.T_INV_LABEL_TEMP) Then
        'ラベル一時テーブルが存在しなかった
        'ラベル一時テーブルを作成する
        Dim dicReplaceLabelTemp As Dictionary
        Set dicReplaceLabelTemp = New Dictionary
        dicReplaceLabelTemp.Add 0, INV_CONST.T_INV_LABEL_TEMP
        dicReplaceLabelTemp.Add 1, clsEnumfrmBIN.INVMasterTana(F_INV_Tana_Local_Text_IMT)
        dicReplaceLabelTemp.Add 2, clsEnumfrmBIN.INVMasterParts(F_Tehai_Code_IMPrt)
        dicReplaceLabelTemp.Add 3, clsEnumfrmBIN.INVMasterParts(F_Label_Name_1_IMPrt)
        dicReplaceLabelTemp.Add 4, clsEnumfrmBIN.INVMasterParts(F_Label_Name_2_IMPrt)
        dicReplaceLabelTemp.Add 5, clsEnumfrmBIN.INVMasterParts(F_Label_Remark_1_IMPrt)
        dicReplaceLabelTemp.Add 6, clsEnumfrmBIN.INVMasterParts(F_Label_Remark_2_IMPrt)
        dicReplaceLabelTemp.Add 7, PublicConst.INPUT_DATE
        dicReplaceLabelTemp.Add 8, INV_CONST.F_INV_LABEL_TEMP_TEHAICODE_LENGTH
        dicReplaceLabelTemp.Add 9, INV_CONST.F_INV_LABEL_TEMP_ORDERNUM
        dicReplaceLabelTemp.Add 10, INV_CONST.F_INV_LABEL_TEMP_SAVEPOINT
        dicReplaceLabelTemp.Add 11, INV_CONST.F_INV_LABEL_TEMP_FORMSTARTTIME
        dicReplaceLabelTemp.Add 12, clsEnumfrmBIN.INVMasterParts(F_Store_Code_IMPrt)
        'Replace実行、SQL設定
        clsADOLabelTemp.SQL = clsSQLBc.ReplaceParm(INV_CONST.SQL_INV_CREATE_LABEL_TEMP_TABLE, dicReplaceLabelTemp)
        'Writeフラグ立てる
        clsADOLabelTemp.ConnectMode = clsADOLabelTemp.ConnectMode Or adModeWrite
        'SQL実行
        Dim isCollect As Boolean
        isCollect = clsADOLabelTemp.Do_SQL_with_NO_Transaction
        'Writeフラグ下げる
        clsADOLabelTemp.ConnectMode = clsADOLabelTemp.ConnectMode And Not adModeWrite
        If Not isCollect Then
            'SQL実行失敗
            DebugMsgWithTime "RecreateLabelTempTable : do sql fail..."
            MsgBox "RecreateLabelTempTableでSQLの実行に失敗しました"
            RecreateLabelTempTable = False
            GoTo CloseAndExit
        End If
    End If
    'メンバ変数のRecordSetに一時テーブルの内容を反映する
    If rsLabelTemp Is Nothing Then
        Set rsLabelTemp = New ADODB.Recordset
    End If
    If rsLabelTemp.State And ObjectStateEnum.adStateOpen Then
        '接続が開いていたら閉じる
        rsLabelTemp.Close
    End If
    rsLabelTemp.ActiveConnection = clsADOLabelTemp.ConnectionString
    rsLabelTemp.Source = "SELECT * FROM " & INV_CONST.T_INV_LABEL_TEMP
    rsLabelTemp.CursorLocation = adUseClient
    rsLabelTemp.CursorType = adOpenStatic
    rsLabelTemp.LockType = adLockBatchOptimistic
    rsLabelTemp.Open , , , , adCmdText
    clsADOLabelTemp.CloseClassConnection
    'オーダーNoテキストボックスのMaxlengthにRSのDefinedSizeをセットしてやる
    txtBox_OrderNumber.MaxLength = rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_ORDERNUM).DefinedSize
    DebugMsgWithTime "RecreateLabelTempTable: Recreate Label Temp Table Success"
    RecreateLabelTempTable = True
    GoTo CloseAndExit
    Exit Function
ErrorCatch:
    DebugMsgWithTime "RecreateLabelTempTable code: " & Err.Number & " Description: " & Err.Description
    RecreateLabelTempTable = False
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOLabelTemp Is Nothing Then
        clsADOLabelTemp.CloseClassConnection
        Set clsADOLabelTemp = Nothing
    End If
    Exit Function
End Function
'''現在のRSのデータをラベルテーブルに追加する
Private Sub AddNewRStoLabelTemp()
    On Error GoTo ErrorCatch
    If Not rsLabelTemp.State And ObjectStateEnum.adStateOpen Then
        '接続していなかったら接続する
        rsLabelTemp.Open
    End If
    '枚数選択されてなかったら抜ける
    If cmbBox_AddLabelCount.ListIndex = -1 Then
        MsgBox "追加する枚数を選択して下さい"
        cmbBox_AddLabelCount.SetFocus
        Exit Sub
    End If
    '枚数コンボボックス分追加する
    Dim longAddCount As Long
    For longAddCount = 1 To cmbBox_AddLabelCount.List(cmbBox_AddLabelCount.ListIndex)
        '新規レコードを追加する
        rsLabelTemp.AddNew
        Dim varKeyobjDic As Variant
        'dicObjtoFieldをループし、rsLabelTempにデータを設定する
        If dicObjNameToFieldName.Exists(Empty) Then
            dicObjNameToFieldName.Remove Empty
        End If
        For Each varKeyobjDic In dicObjNameToFieldName
            Select Case True
            '暫定対応、ラベルTableに追加しない項目もフォームに表示するようになったため
            'Labelコントロールの場合は何もしないで抜ける
            Case TypeName(Me.Controls(varKeyobjDic)) = "Label"
                'ラベルコントロールの時は何もしない
            Case IsNull(clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(varKeyobjDic)).Value)
                'Nullだった場合
                'とりあえず空文字にする
                rsLabelTemp.Fields(dicObjNameToFieldName(varKeyobjDic)).Value = ""
            Case Else
                'データがあった場合
                'RSのデータをそのまま適用する
                rsLabelTemp.Fields(dicObjNameToFieldName(varKeyobjDic)).Value = _
                clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(varKeyobjDic)).Value
            End Select
        Next varKeyobjDic
        '今回のフォームスタート時間をFormStartTimeとして入力
        rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_FORMSTARTTIME).Value = strStartTime
        'InputDateは現在時刻
        rsLabelTemp.Fields(PublicConst.INPUT_DATE).Value = GetLocalTimeWithMilliSec
        '手配コードの文字列数をセット
        rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_TEHAICODE_LENGTH).Value = Len(Trim(rsLabelTemp.Fields(clsEnumfrmBIN.INVMasterParts(F_Tehai_Code_IMPrt)).Value))
        'オーダーNoをセット、先頭からフィールドサイズ分まで
        rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_ORDERNUM).Value = Mid(txtBox_OrderNumber.Text, 1, rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_ORDERNUM).DefinedSize)
        'SavePointをセット、先頭からフィールドサイズ分まで
        rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_SAVEPOINT).Value = Mid(strSavePoint, 1, rsLabelTemp.Fields(INV_CONST.F_INV_LABEL_TEMP_SAVEPOINT).DefinedSize)
        'UpdateでローカルのRSを確定する
        rsLabelTemp.Update
        'InputDateをずらすためSleep 1 (0.001)
        Sleep 1
    Next longAddCount
    'rsLabelのFilterをPendingRecords、変更を未送信に設定し、UpdateBatchをかけ、DBに反映する
    rsLabelTemp.Filter = adFilterNone
    rsLabelTemp.Filter = adFilterPendingRecords
    rsLabelTemp.UpdateBatch adAffectGroup
    rsLabelTemp.Filter = adFilterNone
    If rsLabelTemp.Status And ADODB.RecordStatusEnum.adRecUnmodified Then
        MsgBox "正常に " & cmbBox_AddLabelCount.List(cmbBox_AddLabelCount.ListIndex) & " 枚一時テーブルに追加されました"
        '次の連続指定のために棚番テキストボックスにSetFocus
        txtBox_F_INV_Tana_Local_Text.SetFocus
    End If
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "AddNewRStoLabelTemp code: " & Err.Number & " Description: " & Err.Description
    MsgBox "ラベル印刷用一時テーブル登録時にエラーが発生したため、今回の登録はキャンセルされました"
    rsLabelTemp.CancelUpdate
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''Label Tempテーブルから差し込み印刷を実行する
'''args
'''strargTemplateWordFile               差し込み印刷フィールド設定済みテンプレート文書
'''Optional strargPlaneDocTemplete      プリンタ変更が必要なテンプレート等、Applicaionイベントを使用したい場合に指定する、空のファイル
'''                                     空ファイルの要件は､イベントに必要なmodLabel_BINとclsWordAppEventsクラスを実装していること
Private Sub MailMergeDocCreate(strargMailMergeTemplateFile As String, Optional strargPlaneDocTemplete As String)
    On Error GoTo ErrorCatch
    'テンプレート文書の存在確認
    Dim fsoMailMerge As FileSystemObject
    Set fsoMailMerge = New FileSystemObject
    'clsADOはデフォルトのDBディレクトリを取得する位にしか使わないので単独で作成
    Dim clsADOMailMerge As clsADOHandle
    Set clsADOMailMerge = CreateclsADOHandleInstance
    'DBPathをデフォルトに
    clsADOMailMerge.SetDBPathandFilenameDefault
    'DBFileNameをTempDBに
    clsADOMailMerge.DBFileName = PublicConst.TEMP_DB_FILENAME
    If Not fsoMailMerge.FileExists(strargMailMergeTemplateFile) Then
        'ファイルが存在しなかった
        MsgBox "差し込み印刷用のテンプレートファイルが見つかりませんでした"
        GoTo CloseAndExit
    End If
    'PlaneDocが指定されている場合はそちらのファイルの存在確認も
    If strargPlaneDocTemplete <> "" Then
        If Not fsoMailMerge.FileExists(strargPlaneDocTemplete) Then
            MsgBox "差し込み印刷用のイベント定義ファイルが見つかりませんでした。"
            GoTo CloseAndExit
        End If
    End If
    'T_LABEL_TEMPの存在確認
    If Not clsADOMailMerge.IsTableExists(INV_CONST.T_INV_LABEL_TEMP) Then
        DebugMsgWithTime "MailMergeDocCreate : T_LABEL_TEMP not exists"
        MsgBox "印刷用のデータが登録されていませんでした"
        GoTo CloseAndExit
    End If
    'SavePointの確認、選択を行う
    Dim isCollect As Boolean
    isCollect = CheckandSelectSavePoint
    If Not isCollect Then
        DebugMsgWithTime "MailMergeDocCreate : select save point fail"
'        MsgBox "印刷リストの選択でエラーが発生しました"
        GoTo CloseAndExit
    End If
    'SQLを設定
    Dim strSQL As String
    Dim dicReplace As Dictionary
    Set dicReplace = New Dictionary
    dicReplace.RemoveAll
'    'BinLabel MailMerge用基礎データ取得、選択結果をWhere条件で適用
    dicReplace.Add 0, INV_CONST.T_INV_LABEL_TEMP
    dicReplace.Add 1, clsSQLBc.ReturnMailMergeWhere(frmBinLabel.varstrarrSelectedSavepoint)
    dicReplace.Add 2, INV_CONST.T_INV_SELECT_TEMP
    '255文字を超えてるとMailMergeのSQLとしてはだめらしいので、一旦作業用テーブルに書き出す
    'Writeフラグ上げる
    clsADOMailMerge.ConnectMode = clsADOMailMerge.ConnectMode Or adModeWrite
    'SelectTempテーブル削除
    isCollect = clsADOMailMerge.DropTable(INV_CONST.T_INV_SELECT_TEMP)
    If Not isCollect Then
        MsgBox "中間種作業用テーブルの削除に失敗しました"
        GoTo CloseAndExit
    End If
    clsADOMailMerge.SQL = clsSQLBc.ReplaceParm(INV_CONST.SQL_LABEL_MAILMERGE_DEFAULT, dicReplace)
    '昇順ソートオプションが指定されていたらSQLに追記してやる
    If chkBox_SortByLocationASC.Value Then
        '棚番昇順ソート指定がなされていた
        clsADOMailMerge.SQL = clsADOMailMerge.SQL & vbCrLf & " ORDER BY [" & clsEnumfrmBIN.INVMasterTana(F_INV_Tana_Local_Text_IMT) & "] ASC" & _
        ",[" & INV_CONST.F_INV_LABEL_TEMP_ORDERNUM & "] ASC"
    End If
    'Seletc INTO 実行
    isCollect = clsADOMailMerge.Do_SQL_with_NO_Transaction
    'writeフラグ下げる
    clsADOMailMerge.ConnectMode = clsADOMailMerge.ConnectMode And Not adModeWrite
    If Not isCollect Then
        MsgBox "抽出結果を中間作業用テーブルに登録する際にエラーが発生しました"
        GoTo CloseAndExit
    End If
    'MailMergeはSelectedTableの全選択オンリーで
    strSQL = "SELECT * FROM [" & INV_CONST.T_INV_SELECT_TEMP & "]"
    'wordDocumentsを得る
#If RefWord Then
    'wordの参照設定がされている場合
    Dim objWord As Word.Application
    Set objWord = New Word.Application
#Else
    Dim objWord As Object
    Set objWord = CreateObject("Word.Application")
#End If
'    Dim docTemplateMailMerge As Word.Document
    Dim docTemplateMailMerge As Object
    'ラベルプリント用テンプレートを開く
    Set docTemplateMailMerge = objWord.Documents.Open(Filename:=strargMailMergeTemplateFile)
    With docTemplateMailMerge.MailMerge
        'データソースを開く
        .OpenDataSource Name:=fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, PublicConst.TEMP_DB_FILENAME), ReadOnly:=True, sqlstatement:=strSQL
        '結果は新規ドキュメントへ
        .Destination = 0                'wdSendToNewDocument
        '差し込み印刷実行
        .Execute
    End With
    '差し込み印刷の結果のDocumentを取得
#If RefWord Then
    'Wordが参照設定されている場合
    Dim docNewMailMerge As Word.Document
#Else
    'wordが参照設定されていない場合
    Dim docNewMailMerge As Object
#End If
    Set docNewMailMerge = objWord.ActiveDocument
    'オリジナルのDocumentは保存せずに閉じる
    docTemplateMailMerge.Close savechanges:=False
    'ここから先の処理はApplicationイベントを処理する必要のあるファイルのみ
    If Not strargPlaneDocTemplete = "" Then
        '差し込み結果を一時保存するためのファイル名を取得
        Dim strTempMailmergeFullPath As String
        strTempMailmergeFullPath = fsoMailMerge.BuildPath(clsADOMailMerge.DBPath, GetTimeForFileNameWithMilliSec & "_Local.docx")
        '差し込み結果を一時ファイルに保存、保存形式はdoc Xmlフォーマット(デフォルトを明示的に指定)
        docNewMailMerge.SaveAs2 Filename:=strTempMailmergeFullPath, FileFormat:=16              'wdFormatDocumentDefault 16
        '保存が終わったらDocumentを閉じる
        docNewMailMerge.Close savechanges:=False
        'ラベル印刷用Plane文書を開き、Documentオブジェクトを得る
#If RefWord Then
        'Word参照設定がされている場合
        Dim docLabelPlane As Word.Document
#Else
        '参照設定なしの場合
        Dim docLabelPlane As Object
#End If
        'Label用Plane文書をテンプレートして新規文書を開く
        objWord.Documents.Add Template:=strargPlaneDocTemplete
        '新規文書のDobumentオブジェクトを得る
        Set docLabelPlane = objWord.ActiveDocument
        'Applicatoinイベントハンドラ用に、objWordのApplication参照をセットしてやる
        objWord.Run "modLabel_BIN.SetAppRefForEvent", objWord
        '開いたPlane文書の先頭に差し込み結果をインポートする
        docLabelPlane.Range(0, 0).InsertFile Filename:=strTempMailmergeFullPath, link:=False, attachment:=False
        'インポート完了したら一時保存した差し込み結果ファイルを削除する
        Kill strTempMailmergeFullPath
    End If
    'ここから共通処理
    objWord.Visible = True
    'テーブル削除はプロシージャ分離する
    DeleteCompDataFromLabelTemp
    '印刷終わったらSavePointo消去する
    strSavePoint = ""
    lbl_SavePointName.Caption = ""
    FormCommon.strSavePointName = ""
    ForceForeground objWord.Windows(1).hwnd
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "btnCreateMailmergeDoc_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOMailMerge Is Nothing Then
        clsADOMailMerge.CloseClassConnection
        Set clsADOMailMerge = Nothing
    End If
    If Not objWord Is Nothing Then
'        objWord.Quit
'        Set objWord = Nothing
    End If
    Set fsoMailMerge = Nothing
    Exit Sub
End Sub
'''印刷する際の識別名を選択する
'''Return bool  成功したらTrue、それ以外はFalse
Private Function CheckandSelectSavePoint() As Boolean
    On Error GoTo ErrorCatch
    'ここのみで使用するAdoなので個別宣言する
    Dim clsadoSavePoint As clsADOHandle
    Set clsadoSavePoint = CreateclsADOHandleInstance
    'デフォルトディレクトリへ
    clsadoSavePoint.SetDBPathandFilenameDefault
    'DBファイル名のみ一時DBの物へ
    clsadoSavePoint.DBFileName = PublicConst.TEMP_DB_FILENAME
    'LabelTempテーブルの存在をチェックし、無ければ作成する
    Dim isCollect As Boolean
    If Not clsadoSavePoint.IsTableExists(INV_CONST.T_INV_LABEL_TEMP) Then
        'LabelTempTableが存在しなかった
        isCollect = RecreateLabelTempTable
        If Not isCollect Then
            DebugMsgWithTime "CheckandSelectSavePoint : Create LabelTempTable Fail"
            MsgBox "印刷情報格納一時テーブルの作成に失敗しました"
            GoTo CloseAndExit
        End If
    End If
    'SQL設定、Group By SavePoint,InputDate Oeder by inputdate desc,SavePoint asc
    Dim dicReplace As Dictionary
    Set dicReplace = New Dictionary
    dicReplace.RemoveAll
    dicReplace.Add 0, INV_CONST.F_INV_LABEL_TEMP_SAVEPOINT
    dicReplace.Add 1, INV_CONST.F_INV_LABEL_TEMP_FORMSTARTTIME
    dicReplace.Add 2, INV_CONST.T_INV_LABEL_TEMP
    dicReplace.Add 3, INV_CONST.F_INV_LABEL_TEMP_SAVE_FRENDLYNAME
    dicReplace.Add 4, INV_CONST.F_INV_LABEL_TEMP_FRMSTART_FRENDLYNAME
    clsadoSavePoint.SQL = clsSQLBc.ReplaceParm(INV_CONST.SQL_SELECT_SAVEPOINT, dicReplace)
    'SQL実行
    isCollect = clsadoSavePoint.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        'SQL実行失敗
        MsgBox "CheckandSetSavePoint : ラベル一時テーブルの識別番号一覧読み取りに失敗しました"
        CheckandSelectSavePoint = False
        GoTo CloseAndExit
    End If
    'RecordCountにより処理を分岐
    Select Case clsadoSavePoint.RecordCount
    Case 0
        'レコード無しの場合
        MsgBox "ラベル一時テーブルにデータが見つかりませんでした"
        CheckandSelectSavePoint = False
        GoTo CloseAndExit
    Case 1
        '1個の場合は選択画面を出さずに現在の項目をセットしてやる
        ReDim varstrarrSelectedSavepoint(0, 1)
        'rs.MoveFirstする
        clsadoSavePoint.RS.MoveFirst
        varstrarrSelectedSavepoint(0, 0) = CStr(clsadoSavePoint.RS.Fields(INV_CONST.F_INV_LABEL_TEMP_SAVE_FRENDLYNAME).Value)
        varstrarrSelectedSavepoint(0, 1) = CStr(clsadoSavePoint.RS.Fields(INV_CONST.F_INV_LABEL_TEMP_FRMSTART_FRENDLYNAME).Value)
        CheckandSelectSavePoint = True
        GoTo CloseAndExit
    Case Is >= 2
        '複数個のSavePointが見つかった場合は選択画面を設定、表示する
        Load frmSelectSavePoint
        frmSelectSavePoint.lstBoxSavePoint.List = clsadoSavePoint.RS_Array
        frmSelectSavePoint.lstBoxSavePoint.ColumnCount = 2
        'Modalで選択画面を表示する(完了したら表示先のフォームで自身Unloadして処理が戻ってくるはず)
        frmSelectSavePoint.Show vbModal
        If IsEmpty(varstrarrSelectedSavepoint) Then
            '結果VariantにEmptyがセットされていた、選択画面で何か失敗したっぽい
'            MsgBox "リスト識別名選択画面でエラーが発生しました"
            CheckandSelectSavePoint = False
            GoTo CloseAndExit
        End If
        'ここまで抜けてきたらとりあえず選択は上手くいったと認識
        CheckandSelectSavePoint = True
        GoTo CloseAndExit
    End Select
ErrorCatch:
    CheckandSelectSavePoint = False
    DebugMsgWithTime "CheckandSelectSavePoint code: " & Err.Number & " Description " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsadoSavePoint Is Nothing Then
        '接続切断とインスタンス破棄
        clsadoSavePoint.CloseClassConnection
        Set clsadoSavePoint = Nothing
    End If
    Exit Function
End Function
'''SavePointをセットする
Private Sub SetSavePoint()
    Select Case True
    Case strSavePoint = ""
        'strSavePointが空文字だったということは今回初実行なので、識別名を入力してもらう
        'frmSetSavePoint 表示、終わったら勝手に向こうでUnloadして処理が戻ってくるはず
        frmSetSavePoint.Show vbModal
        If FormCommon.strSavePointName = "" Then
            '空文字が返ってきた、キャンセルされたか空文字にされたかどっちか
            'GetLocalTimeで良いと思う
            FormCommon.strSavePointName = GetLocalTimeWithMilliSec
        End If
        '結果をローカル変数に反映
        strSavePoint = FormCommon.strSavePointName
        '23文字超えてたら切り下げる
        If Len(strSavePoint) > 23 Then
            strSavePoint = Mid(strSavePoint, 1, 23)
        End If
        'SavePointラベルに値をセット
        lbl_SavePointName.Caption = strSavePoint
        'strStartTimeに時間をセット
        strStartTime = GetLocalTimeWithMilliSec
        Exit Sub
    End Select
End Sub
'''印刷が完了したLabelTempのデータを削除する
'''レコードカウントが0になったらテーブルを消去する
Private Sub DeleteCompDataFromLabelTemp()
    On Error GoTo ErrorCatch
    '個別にConnection張りたいので、独立してclsADOを定義
    Dim clsAdoDeleteComp As clsADOHandle
    Set clsAdoDeleteComp = CreateclsADOHandleInstance
    'DBPath,DBFilenameを一時DBの物へ
    clsAdoDeleteComp.SetDBPathandFilenameDefault
    clsAdoDeleteComp.DBFileName = PublicConst.TEMP_DB_FILENAME
    Dim dicRepladeDelete As Dictionary
    Set dicRepladeDelete = New Dictionary
    dicRepladeDelete.RemoveAll
    dicRepladeDelete.Add 0, INV_CONST.T_INV_LABEL_TEMP
    dicRepladeDelete.Add 1, clsSQLBc.ReturnMailMergeWhere(varstrarrSelectedSavepoint)
    'SQL設定
    clsAdoDeleteComp.SQL = clsSQLBc.ReplaceParm(SQL_DELET_COMP_DATA_FROM_LABEL_TEMP, dicRepladeDelete)
    'Writeフラグ上げる
    clsAdoDeleteComp.ConnectMode = clsAdoDeleteComp.ConnectMode Or adModeWrite
    'SQL実行
    Dim isCollect As Boolean
    isCollect = clsAdoDeleteComp.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        MsgBox "印刷完了データの削除に失敗しました"
        GoTo CloseAndExit
    End If
    '残りデータ0だったらテーブル削除
    '一旦LabelTempの全データ取得
    clsAdoDeleteComp.SQL = "SELECT * FROM " & INV_CONST.T_INV_LABEL_TEMP
    'SQL実行
    isCollect = clsAdoDeleteComp.Do_SQL_with_NO_Transaction
    If clsAdoDeleteComp.RecordCount = 0 Then
        'RecordCount0だったらテーブル削除
        isCollect = clsAdoDeleteComp.DropTable(INV_CONST.T_INV_LABEL_TEMP)
    End If
    If Not isCollect Then
        MsgBox "一時テーブルの削除に失敗しました"
        GoTo CloseAndExit
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "DeleteCompDataFromLabelTemp code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    If Not clsAdoDeleteComp Is Nothing Then
        'Writeフラグ下げる
        clsAdoDeleteComp.ConnectMode = clsAdoDeleteComp.ConnectMode And Not adModeWrite
        'コネクション切断
        clsAdoDeleteComp.CloseClassConnection
        Set clsAdoDeleteComp = Nothing
    End If
    Exit Sub
End Sub
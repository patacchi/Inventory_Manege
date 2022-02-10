VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBinLabel 
   Caption         =   "BINカードラベル印刷項目編集画面"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610.001
   OleObjectBlob   =   "frmBinLabel.frx":0000
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
Private rsfrmBIN As ADODB.Recordset
Private confrmBIN As ADODB.Connection
Private StopEvents As Boolean
'定数
Private Const MAX_LABEL_TEXT_LENGTH As Long = 18
Private Const TXTBOX_BACKCOLORE_EDITABLE As Long = &HC0FFC0         '薄い緑
Private Const TXTBOX_BACKCOLORE_NORMAL As Long = &H80000005         'ウィンドウの背景
'------------------------------------------------------------------------------------------------------
'SQL
Private Const SQL_BIN_LABEL_DEFAULT_DATA As String = "SELECT TDBPrt.F_INV_Tehai_ID,TDBTana.F_INV_Tana_ID,TDBTana.F_INV_Tana_Local_Text as F_INV_Tana_Local_Text,TDBPrt.F_INV_Tehai_Code as F_INV_Tehai_Code,TDBPrt.F_INV_Label_Name_1 as F_INV_Label_Name_1,TDBPrt.F_INV_Label_Name_2 as F_INV_Label_Name_2,TDBPrt.F_INV_Label_Remark_1 as F_INV_Label_Remark_1,TDBPrt.F_INV_Label_Remark_2 as F_INV_Label_Remark_2" & vbCrLf & _
"FROM T_INV_M_Parts AS TDBPrt " & vbCrLf & _
"    INNER JOIN T_INV_M_Tana as TDBTana " & vbCrLf & _
"    ON TDBPrt.F_INV_Tana_ID = TDBTana.F_INV_Tana_ID"
'------------------------------------------------------------------------------------------------------
'イベント
'Form Initial
Private Sub UserForm_Initialize()
    'フォーム初期化時
    ConstRuctor
End Sub
'Form Terminate
Private Sub UserForm_Terminate()
    Destructor
End Sub
'Click
Private Sub btnMovePrevious_Click()
    '前へ戻る
    MoveRecord vbKeyLeft
End Sub
Private Sub btnMoveNext_Click()
    '次へ進む
    MoveRecord vbKeyRight
End Sub
'編集制限解除
Private Sub btnEnableEdit_Click()
    SwitchtBoxEditmode True
End Sub
'最終的にDBにUpdateする
Private Sub btnDoUpdate_Click()
    DoUpdateBatch
End Sub
'Change
'RSに値セットするテキストボックス
'棚番
Private Sub txtBox_F_INV_Tana_Local_Text_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
'    If Len(ActiveControl.Text) > MAX_LABEL_TEXT_LENGTH Then
'        MsgBox "設定可能な文字数" & MAX_LABEL_TEXT_LENGTH & " を超えています。"
'        Exit Sub
'    End If
    'Updateメソッドへ
    UpdateRSFromContrl ActiveControl
End Sub
'品名1
Private Sub txtBox_F_INV_Label_Name_1_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
'    If Len(ActiveControl.Text) > MAX_LABEL_TEXT_LENGTH Then
'        MsgBox "設定可能な文字数" & MAX_LABEL_TEXT_LENGTH & " を超えています。"
'        Exit Sub
'    End If
    'Updateメソッドへ
    UpdateRSFromContrl ActiveControl
End Sub
'品名2
Private Sub txtBox_F_INV_Label_Name_2_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
'    If Len(ActiveControl.Text) > MAX_LABEL_TEXT_LENGTH Then
'        MsgBox "設定可能な文字数" & MAX_LABEL_TEXT_LENGTH & " を超えています。"
'        Exit Sub
'    End If
    'Updateメソッドへ
    UpdateRSFromContrl ActiveControl
End Sub
'備考1
Private Sub txtBox_F_INV_Label_Remark_1_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
'    If Len(ActiveControl.Text) > MAX_LABEL_TEXT_LENGTH Then
'        MsgBox "設定可能な文字数" & MAX_LABEL_TEXT_LENGTH & " を超えています。"
'        Exit Sub
'    End If
    'Updateメソッドへ
    UpdateRSFromContrl ActiveControl
End Sub
'備考2
Private Sub txtBox_F_INV_Label_Remark_2_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
'    If Len(ActiveControl.Text) > MAX_LABEL_TEXT_LENGTH Then
'        MsgBox "設定可能な文字数" & MAX_LABEL_TEXT_LENGTH & " を超えています。"
'        Exit Sub
'    End If
    'Updateメソッドへ
    UpdateRSFromContrl ActiveControl
End Sub
'手配コードフィルタ
Private Sub txtBox_Filter_Tehai_Code_Change()
    If StopEvents Then
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    If Len(ActiveControl.Text) >= 1 Then
        ActiveControl.Text = UCase(ActiveControl.Text)
    End If
    SetFilter ActiveControl
End Sub
'棚番フィルター
Private Sub txtBox_Filter_Local_Tana_Change()
    If StopEvents Then
        Exit Sub
    End If
    'イベント停止
    StopEvents = True
    If Len(ActiveControl.Text) >= 1 Then
        ActiveControl.Text = UCase(ActiveControl.Text)
    End If
    SetFilter ActiveControl
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
    If rsfrmBIN Is Nothing Then
        Set rsfrmBIN = New ADODB.Recordset
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
    'RSのデータを取得する
    GetValuFromRS
#If DebugDB Then
    MsgBox "DebugDB有効"
#End If
End Sub
'''デストラクタ
Private Sub Destructor()
    'メンバ変数の解放、特に接続が関連しているものは重点的に
    If Not clsADOfrmBIN Is Nothing Then
        clsADOfrmBIN.CloseClassConnection
        Set clsADOfrmBIN = Nothing
    End If
    If Not rsfrmBIN Is Nothing Then
        rsfrmBIN.ActiveConnection.Close
'        rsfrmBIN.Close
        Set rsfrmBIN = Nothing
    End If
    If Not confrmBIN Is Nothing Then
        If confrmBIN.State And ObjectStateEnum.adStateOpen Then
            '接続していたら閉じる
            confrmBIN.Close
        End If
        Set confrmBIN = Nothing
    End If
End Sub
'''メンバ変数のRecordSetに初期データを設定する
'''
Private Sub SetDefaultValuetoRS()
    '最初にclsadoのDBPathとDBFilnameをデフォルトに
    clsADOfrmBIN.SetDBPathandFilenameDefault
    'もし接続されていたら切断する
    If rsfrmBIN.State And ObjectStateEnum.adStateOpen Then
        rsfrmBIN.Close
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
    rsfrmBIN.LockType = adLockBatchOptimistic
    rsfrmBIN.CursorType = adOpenStatic
    'rsのSourceにSQL設定(後でパラメータ対応する)
    rsfrmBIN.Source = SQL_BIN_LABEL_DEFAULT_DATA
    'rsのActiveConnectionにConnectionオブジェクト指定
    Set rsfrmBIN.ActiveConnection = confrmBIN
    'rsオープン
    rsfrmBIN.Open , , , , CommandTypeEnum.adCmdText
    '以下は正常に動く
    '更新に必要なキー列の情報が～・・・→両方のテーブルの主キーをSELECTのフィールドに含めると解決
'    rsfrmBIN.Fields("F_INV_Label_Name_2").Value = "InputTest"
'    rsfrmBIN.Fields("F_INV_Tana_Local_Text").Value = "K23 A01"
'    rsfrmBIN.Update
'    rsfrmBIN.UpdateBatch
    DebugMsgWithTime "Default Data count: " & rsfrmBIN.RecordCount
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
End Sub
'cidObjToFieldにあるコントロールの値をすべて消去する
Private Sub ClearAllContents()
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
End Sub
'RSから値をとってくる
Private Sub GetValuFromRS()
    On Error GoTo ErrorCatch
    If rsfrmBIN.EOF And rsfrmBIN.BOF Then
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
        Case IsNull(rsfrmBIN.Fields(dicObjNameToFieldName(varKeyobjDic)).Value)
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
                Me.Controls(varKeyobjDic).Text = rsfrmBIN.Fields(dicObjNameToFieldName(varKeyobjDic)).Value
            Case "Label"
                'ラベル
                Me.Controls(varKeyobjDic).Caption = rsfrmBIN.Fields(dicObjNameToFieldName(varKeyobjDic)).Value
            End Select
        End Select
    Next varKeyobjDic
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "GetValuFromRS code: " & err.Number & " Description: " & err.Description
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
    If rsfrmBIN.BOF And rsfrmBIN.EOF Then
        'BOFとEOF両方立ってたら抜ける
    End If
    'イベント停止する
    StopEvents = True
    Select Case intargKeyCode
    Case vbKeyRight
        '右、次へ
        rsfrmBIN.MoveNext
        If rsfrmBIN.EOF Then
            MsgBox "現在のレコードが最終レコードです"
            rsfrmBIN.MovePrevious
        End If
    Case vbKeyLeft
        '左、前へ
        rsfrmBIN.MovePrevious
        If rsfrmBIN.BOF Then
            MsgBox "現在のレコードが先頭レコードです"
            rsfrmBIN.MoveNext
        End If
    End Select
    '値の取得をする
    GetValuFromRS
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "MoveRecord code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'フィルタテキストボックスでChangeイベントが発生したらRSにFilter設定してやる
Private Sub SetFilter(ByRef argCtrl As Control)
    If rsfrmBIN.BOF And rsfrmBIN.EOF Then
        'RSに中身が無かったら抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    Select Case argCtrl.Text
    Case ""
        '空白だったら、FilterにadFilterNonをセットしてフィルタをクリアする
        rsfrmBIN.Filter = adFilterNone
        '値取得する
        GetValuFromRS
    Case Else
        '何かしら文字列が入ってたら、Like ～%といった感じで前方一致で条件を組む
        Dim strFilter(3) As String
        Select Case argCtrl.Name
        Case txtBox_Filter_Local_Tana.Name
            '棚番だった場合
            strFilter(0) = dicObjNameToFieldName(txtBox_F_INV_Tana_Local_Text.Name)
        Case txtBox_Filter_Tehai_Code.Name
            '手配コードだった場合
            strFilter(0) = dicObjNameToFieldName(txtBox_F_INV_Tehai_Code.Name)
        End Select
        '共通部分を埋めていく
        strFilter(1) = " LIKE '"
        strFilter(2) = argCtrl.Text
        '最後にワイルドカード付与
        strFilter(3) = "%'"
        'Filterセット
        rsfrmBIN.Filter = Join(strFilter, "")
        '値取得する
        GetValuFromRS
    End Select
    'レコードが0だったら報告する
    If rsfrmBIN.BOF And rsfrmBIN.EOF Then
        MsgBox "現在の指定条件では該当するレコードがありません"
        '一旦フィルタ解除する
        rsfrmBIN.Filter = adFilterNone
        'テキストボックスの文字数により処理を分岐
        'イベント再開
        StopEvents = False
        Select Case Len(argCtrl.Text)
        Case Is = 1
            '1文字目でダメだったらテキストを全消去
            argCtrl.Text = ""
        Case Is >= 2
            '2文字以上あったらフィルタ文字列1文字減らして再セット
            argCtrl.Text = Mid(argCtrl.Text, 1, Len(argCtrl.Text) - 1)
        End Select
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SetFilter code: " & err.Number & " Description: " & err.Description
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
        btnDoUpdate.Enabled = True
        'LockedをFalseにして、BackColoreを薄緑にする
        txtBox_F_INV_Tana_Local_Text.Locked = False
        txtBox_F_INV_Tana_Local_Text.BackColor = TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Name_1.Locked = False
        txtBox_F_INV_Label_Name_1.BackColor = TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Name_2.Locked = False
        txtBox_F_INV_Label_Name_2.BackColor = TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Remark_1.Locked = False
        txtBox_F_INV_Label_Remark_1.BackColor = TXTBOX_BACKCOLORE_EDITABLE
        txtBox_F_INV_Label_Remark_2.Locked = False
        txtBox_F_INV_Label_Remark_2.BackColor = TXTBOX_BACKCOLORE_EDITABLE
        '編集可能設定ボタンを無効に
        btnEnableEdit.Enabled = False
    Case False
        '編集不可にするとき
        'UpdateBatckボタンをFalseに
        btnDoUpdate.Enabled = False
        'LockedをTrueにして、BackColoreを標準背景色にする
        txtBox_F_INV_Tana_Local_Text.Locked = True
        txtBox_F_INV_Tana_Local_Text.BackColor = TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Name_1.Locked = True
        txtBox_F_INV_Label_Name_1.BackColor = TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Name_2.Locked = True
        txtBox_F_INV_Label_Name_2.BackColor = TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Remark_1.Locked = True
        txtBox_F_INV_Label_Remark_1.BackColor = TXTBOX_BACKCOLORE_NORMAL
        txtBox_F_INV_Label_Remark_2.Locked = True
        txtBox_F_INV_Label_Remark_2.BackColor = TXTBOX_BACKCOLORE_NORMAL
        '編集可能設定ボタンを有効に
        btnEnableEdit.Enabled = True
    End Select
End Sub
'各コントロールの値をRSにセットする
Private Sub UpdateRSFromContrl(argCtrl As Control)
    On Error GoTo ErrorCatch
    If Not dicObjNameToFieldName.Exists(argCtrl.Name) Then
        'dicobjToFieldに存在しないコントロール名の場合は抜ける
        Exit Sub
    End If
    Select Case True
    '最初に文字数チェックを行い、オーバーしていたら設定値まで切り下げる
    Case Len(argCtrl.Text) > rsfrmBIN.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize
        '文字数がフィールド設定値オーバー
        MsgBox "入力された文字数が設定の " & rsfrmBIN.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize & " 文字を超えています。"
        argCtrl.Text = Mid(argCtrl.Text, 1, rsfrmBIN.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize)
        GoTo CloseAndExit
    Case IsNull(rsfrmBIN.Fields(dicObjNameToFieldName(argCtrl.Name)).Value), rsfrmBIN.Fields(dicObjNameToFieldName(argCtrl.Name)).Value <> argCtrl.Text
        'RSの値がNullか、引数のコントロールのtextと違っている場合
        'rsに値をセットして、Updateまでする（DBに反映するにはUpdateBatchしないとダメ）
        rsfrmBIN.Fields(dicObjNameToFieldName(argCtrl.Name)).Value = _
        argCtrl.Text
        rsfrmBIN.Update
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "UpdateRSFromContrl code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'UpdateでRSにコミットされた変更をDBにプッシュする
Private Sub DoUpdateBatch()
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    If rsfrmBIN.Status And adRecModified Then
        rsfrmBIN.UpdateBatch
'        If (rsfrmBIN.Status And ADODB.RecordStatusEnum.adRecUnmodified) Or (rsfrmBIN.Status = ADODB.RecordStatusEnum.adRecOK) Then
        If (rsfrmBIN.Status And ADODB.RecordStatusEnum.adRecUnmodified) Then
            MsgBox "正常に更新されました"
            '編集不可モードへ
            SwitchtBoxEditmode False
            GoTo CloseAndExit
        Else
            MsgBox "更新に失敗した可能性があります RSStasus: " & rsfrmBIN.Status
            GoTo CloseAndExit
        End If
    ElseIf rsfrmBIN.Status And ADODB.RecordStatusEnum.adRecUnmodified Then
        MsgBox "変更点はありませんでした。"
        GoTo CloseAndExit
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "DoUpdateBatch code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBinLabel 
   Caption         =   "BINカードラベル印刷項目編集画面"
   ClientHeight    =   5895
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
Private StopEvents As Boolean
Private UpdateMode As Boolean                                       '編集可能状態になってるときはTrueをセット
Private strStartTime As String
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
    Constructor
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
'変更を破棄
Private Sub btnCancelUpdate_Click()
    CancelUpdateBatch
End Sub
'ラベル一時テーブルに追加する
Private Sub btnAddnewLabelTemp_Click()
    If strStartTime = "" Then
        'フォームスタート時間が空文字だったら今回のフォームでの初回実行とみなす
        'まずはテーブルを作り直す
        Dim isCollect As Boolean
        isCollect = RecreateLabelTempTable
        If Not isCollect Then
            '一時テーブル作成に失敗
            MsgBox "一時テーブル作成に失敗したため、処理を中断します"
            Exit Sub
        End If
        'フォームスタート時間を設定する
        strStartTime = GetLocalTimeWithMilliSec
    End If
    '次にカレントレコードをTempTableに追加する
    AddNewRStoLabelTemp
End Sub
'インクリメンタルリストClick
Private Sub lstBox_Incremental_Click()
    If StopEvents Or UpdateMode Then
        'イベントストップかUpdateModeの時は抜ける
        Exit Sub
    End If
    If clsIncrementalfrmBIN.Incremental_LstBox_Click Then
        'この中に入ってる時点でRSにフィルターが適用されている
        'イベント停止
        StopEvents = True
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
        'イベント再開
        StopEvents = False
    End If
End Sub
'Change
Private Sub txtBox_F_INV_Tehai_Code_Change()
    'イベント停止状態ではなく、更にアップデートモードでもないときにインクリメンタル実行
    If StopEvents Or UpdateMode Then
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'テキストにUcaseかける
    If ActiveControl.TextLength >= 1 Then
        ActiveControl.Text = UCase(ActiveControl.Text)
    End If
    'インクリメンタル実行
    clsIncrementalfrmBIN.Incremental_TextBox_Change
    'イベント再開する
    StopEvents = False
End Sub
'keyup
'インクリメンタルリストKeyUp
Private Sub lstBox_Incremental_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'イベント停止
    StopEvents = True
    'インクリメンタルに丸投げ
    clsIncrementalfrmBIN.Incremental_LstBox_Key_UP KeyCode, Shift
    'イベント再開
    StopEvents = False
End Sub
'mouseup
'インクリメンタルリストMouseUP
Private Sub lstBox_Incremental_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'イベント停止
    StopEvents = True
    'インクリメンタル
    clsIncrementalfrmBIN.Incremental_LstBox_Mouse_UP Button
    'イベント再開
    StopEvents = False
End Sub
'RSに値セットするテキストボックス
'棚番
Private Sub txtBox_F_INV_Tana_Local_Text_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    Select Case UpdateMode
    Case True
        'アップデートモードの時
        UpdateRSFromContrl ActiveControl
        Exit Sub
    Case False
        '検索モード(?)の時
        'イベント停止
        StopEvents = True
        'RSから内容を取得(listboxのClickイベントで呼ばれるはず)されるまでUpdateModeにしてはいけない
        'btnEnableEditをFalseに
        btnEnableEdit.Enabled = False
        'Ucase掛ける
        If ActiveControl.TextLength >= 1 Then
            ActiveControl.Text = UCase(ActiveControl.Text)
        End If
        clsIncrementalfrmBIN.Incremental_TextBox_Change
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
    If UpdateMode Then
        'UpdateModeの時はUpdateメソッドへ
        UpdateRSFromContrl ActiveControl
    End If
End Sub
'品名2
Private Sub txtBox_F_INV_Label_Name_2_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    If UpdateMode Then
        'UpdateModeの時はUpdateメソッドへ
        UpdateRSFromContrl ActiveControl
    End If
End Sub
'備考1
Private Sub txtBox_F_INV_Label_Remark_1_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    If UpdateMode Then
        'UpdateModeの時はUpdateメソッドへ
        UpdateRSFromContrl ActiveControl
    End If
End Sub
'備考2
Private Sub txtBox_F_INV_Label_Remark_2_Change()
    If StopEvents Then
        'イベント停止フラグが立ってたら中止
        Exit Sub
    End If
    If UpdateMode Then
        'UpdateModeの時はUpdateメソッドへ
        UpdateRSFromContrl ActiveControl
    End If
End Sub
'手配コードフィルタ
Private Sub txtBox_Filter_Tehai_Code_Change()
'    If StopEvents Then
'        Exit Sub
'    End If
'    'イベント停止
'    StopEvents = True
'    If Len(ActiveControl.Text) >= 1 Then
'        ActiveControl.Text = UCase(ActiveControl.Text)
'    End If
'    SetFilter ActiveControl
End Sub
'棚番フィルター
Private Sub txtBox_Filter_Local_Tana_Change()
'    If StopEvents Then
'        Exit Sub
'    End If
'    'イベント停止
'    StopEvents = True
'    If Len(ActiveControl.Text) >= 1 Then
'        ActiveControl.Text = UCase(ActiveControl.Text)
'    End If
'    SetFilter ActiveControl
End Sub
'Enter
'棚番テキストボックスEnter
Private Sub txtBox_F_INV_Tana_Local_Text_Enter()
    If StopEvents Or UpdateMode Then
        'StopEvent か UpdateModeだったら抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'インクリメンタル実行、Enter
    clsIncrementalfrmBIN.Incremental_TextBox_Enter txtBox_F_INV_Tana_Local_Text, lstBox_Incremental
    'イベント再開
    StopEvents = False
    Exit Sub
End Sub
Private Sub txtBox_F_INV_Tehai_Code_Enter()
    'イベント停止状態ではなく、更にアップデートモードでもないときにインクリメンタル実行
    If StopEvents Or UpdateMode Then
        'StopEventsかUpdateModeの時は抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'インクリメンタル実行、リストを表示するのが目的
    clsIncrementalfrmBIN.Incremental_TextBox_Enter ActiveControl, lstBox_Incremental
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
Private Sub Constructor()
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
    clsIncrementalfrmBIN.Constructor Me, dicObjNameToFieldName, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc
    'RSのデータを取得する
    'ここでは取得しないで、インクリメンタルサーチに任せる
'    GetValuFromRS
    'イベント再開する
    StopEvents = False
#If DebugDB Then
    MsgBox "DebugDB有効"
#End If
End Sub
'''デストラクタ
Private Sub Destructor()
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
'''
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
    clsADOfrmBIN.RS.Source = SQL_BIN_LABEL_DEFAULT_DATA
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
    DebugMsgWithTime "MoveRecord code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'フィルタテキストボックスでChangeイベントが発生したらRSにFilter設定してやる
Private Sub SetFilter(ByRef argCtrl As Control)
    If clsADOfrmBIN.RS.BOF And clsADOfrmBIN.RS.EOF Then
        'RSに中身が無かったら抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    Select Case argCtrl.Text
    Case ""
        '空白だったら、FilterにadFilterNonをセットしてフィルタをクリアする
        clsADOfrmBIN.RS.Filter = adFilterNone
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
        clsADOfrmBIN.RS.Filter = Join(strFilter, "")
        '値取得する
        GetValuFromRS
    End Select
    'レコードが0だったら報告する
    If clsADOfrmBIN.RS.BOF And clsADOfrmBIN.RS.EOF Then
        MsgBox "現在の指定条件では該当するレコードがありません"
        '一旦フィルタ解除する
        clsADOfrmBIN.RS.Filter = adFilterNone
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
        UpdateMode = True
        btnDoUpdate.Enabled = True
        btnCancelUpdate.Enabled = True
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
        UpdateMode = False
        'UpdateBatckボタンをFalseに
        btnDoUpdate.Enabled = False
        btnCancelUpdate.Enabled = False
        'LockedをTrueにして、BackColoreを標準背景色にする
        '棚番テキストボックスは編集不可モードの時はインクリメンタルに使うのでLockはしない
'        txtBox_F_INV_Tana_Local_Text.Locked = True
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
    Case Len(argCtrl.Text) > clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize
        '文字数がフィールド設定値オーバー
        MsgBox "入力された文字数が設定の " & clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize & " 文字を超えています。"
        argCtrl.Text = Mid(argCtrl.Text, 1, clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).DefinedSize)
        GoTo CloseAndExit
    Case IsNull(clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).Value), clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).Value <> argCtrl.Text
        'RSの値がNullか、引数のコントロールのtextと違っている場合
        'rsに値をセットして、Updateまでする（DBに反映するにはUpdateBatchしないとダメ）
        clsADOfrmBIN.RS.Fields(dicObjNameToFieldName(argCtrl.Name)).Value = _
        argCtrl.Text
        clsADOfrmBIN.RS.Update
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
            '編集不可モードへ
            SwitchtBoxEditmode False
            'RSよりデータを取得する
            GetValuFromRS
            GoTo CloseAndExit
        Else
            MsgBox "更新に失敗した可能性があります RSStasus: " & clsADOfrmBIN.RS.Status
            GoTo CloseAndExit
        End If
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "DoUpdateBatch code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'変更された内容を破棄して元に戻す
Private Sub CancelUpdateBatch()
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
        MsgBox "変更点はありませんでした"
        'フィルタを戻す？
        clsADOfrmBIN.RS.Filter = varOldFilter
        If clsADOfrmBIN.RS.Supports(adBookmark) Then
            'BookMark有効ならBookMarkを戻す
            clsADOfrmBIN.RS.Bookmark = varBookMark
        End If
        '編集不可モードへ
        SwitchtBoxEditmode False
        GoTo CloseAndExit
    Case clsADOfrmBIN.RS.Status And adRecModified
        'Statusが変更有になっている
        'キャンセルしていいか問い合わせ
        Dim longMsgBoxRet As Long
        longMsgBoxRet = MsgBox("内容が変更されています、変更を破棄しても良いですか?", vbYesNo)
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
            MsgBox "変更点を無事に破棄しました。"
            'フィルタを戻す？
            clsADOfrmBIN.RS.Filter = varOldFilter
            If clsADOfrmBIN.RS.Supports(adBookmark) Then
                'BookMark有効ならBookMarkを戻す
                clsADOfrmBIN.RS.Bookmark = varBookMark
            End If
            '編集不可モードへ
            SwitchtBoxEditmode False
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
    DebugMsgWithTime "CancelUpdateBatch code: " & err.Number & " Description: " & err.Description
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
    'まずは既存のラベル一時テーブルを削除
    Dim isCollect As Boolean
    isCollect = clsADOLabelTemp.DropTable(INV_CONST.T_INV_LABEL_TEMP)
    If Not isCollect Then
        DebugMsgWithTime "RecreateLabelTempTable : fail delete already label tamp table"
        MsgBox "ラベル出力一時テーブルの作成に失敗しました"
        RecreateLabelTempTable = False
        GoTo CloseAndExit
        Exit Function
    End If
    'ラベル一時テーブルを作成する
''{0} T_INV_LABEL_TEMP
''{1} F_INV_Tana_Local_Text
''{2} F_INV_Tehai_Code
''{3} F_INV_Label_Name_1
''{4} F_INV_Label_Name_2
''{5} F_INV_Label_Remark_1
''{6} F_INV_Label_Remark_2
''{7} InputDate
'Public Const SQL_INV_CREATE_LABEL_TEMP_TABLE As String = "CREATE TABLE {0} (" & vbCrLf &
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
    'Replace実行、SQL設定
    clsADOLabelTemp.SQL = clsSQLBc.ReplaceParm(INV_CONST.SQL_INV_CREATE_LABEL_TEMP_TABLE, dicReplaceLabelTemp)
    'Writeフラグ立てる
    clsADOLabelTemp.ConnectMode = clsADOLabelTemp.ConnectMode Or adModeWrite
    'SQL実行
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
    DebugMsgWithTime "RecreateLabelTempTable: Recreate Label Temp Table Success"
    RecreateLabelTempTable = True
    GoTo CloseAndExit
    Exit Function
ErrorCatch:
    DebugMsgWithTime "RecreateLabelTempTable code: " & err.Number & " Description: " & err.Description
    RecreateLabelTempTable = False
    GoTo CloseAndExit
CloseAndExit:
    If Not clsADOLabelTemp Is Nothing Then
        clsADOLabelTemp.CloseClassConnection
        Set clsADOLabelTemp = Nothing
    End If
    Exit Function
End Function
'現在のRSのデータをラベルテーブルに追加する
Private Sub AddNewRStoLabelTemp()
    On Error GoTo ErrorCatch
    If Not rsLabelTemp.State And ObjectStateEnum.adStateOpen Then
        '接続していなかったら接続する
        rsLabelTemp.Open
    End If
    '新規レコードを追加する
    rsLabelTemp.AddNew
    Dim varKeyobjDic As Variant
    'dicObjtoFieldをループし、rsLabelTempにデータを設定する
    If dicObjNameToFieldName.Exists(Empty) Then
        dicObjNameToFieldName.Remove Empty
    End If
    For Each varKeyobjDic In dicObjNameToFieldName
        Select Case True
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
    '今回のフォームスタート時間をInputDateとして入力
    rsLabelTemp.Fields(PublicConst.INPUT_DATE).Value = strStartTime
    'UpdateでローカルのRSを確定する
    rsLabelTemp.Update
    'rsLabelのFilterをPendingRecords、変更を未送信に設定し、UpdateBatchをかけ、DBに反映する
    rsLabelTemp.Filter = adFilterNone
    rsLabelTemp.Filter = adFilterPendingRecords
    rsLabelTemp.UpdateBatch adAffectGroup
    rsLabelTemp.Filter = adFilterNone
    If rsLabelTemp.Status And ADODB.RecordStatusEnum.adRecUnmodified Then
        MsgBox "正常に一時テーブルに追加されました"
    End If
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "AddNewRStoLabelTemp code: " & err.Number & " Description: " & err.Description
    MsgBox "ラベル印刷用一時テーブル登録時にエラーが発生したため、今回の登録はキャンセルされました"
    rsLabelTemp.CancelUpdate
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
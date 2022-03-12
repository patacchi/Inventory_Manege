VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegistNewLocation 
   Caption         =   "新規棚番登録画面"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8100
   OleObjectBlob   =   "frmRegistNewLocation.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmRegistNewLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'新規棚番を登録するフォーム
Option Explicit
'クラス共有インスタンス変数
Private clsADONewLocation As clsADOHandle
Private clsEnumNewLocation As clsEnum
Private clsSqlBCNewLocation As clsSQLStringBuilder
Private clsIncrementalNewLocation As clsIncrementalSerch
Private dicObjToFieldName As Dictionary
'メンバ変数
Private StopEvents As Boolean
Private conADONewLocation As ADODB.Connection
'''------------------------------------------------------------------------------------------------------
'SQL定義
'デフォルトデータ取得SQL
'{0}    Tana_ID
'{1}    Tana_System
'{2}    Tana_Local
'{3}    T_M_Tana
'{4}    InputDate
Private Const SQL_DEFAULT_NEW_LOCATION_DATA As String = "SELECT DISTINCT {0},{1},{2},{4} FROM {3} " & vbCrLf & _
                                                        "WHERE {1} IS NOT NULL"
'''------------------------------------------------------------------------------------------------------
'イベント
'Initialize
Private Sub UserForm_Initialize()
    ConstRuctor
End Sub
'Teminate
Private Sub UserForm_Terminate()
    DestRuctor
End Sub
'TextBoxEnter
Private Sub txtBox_F_INV_Tana_Local_Text_Enter()
    If StopEvents Then
        'イベント停止状態だったら抜ける
        Exit Sub
    End If
    'イベント停止する
    StopEvents = True
    'Incremental TextBoxEnter
    clsIncrementalNewLocation.Incremental_TextBox_Enter txtBox_F_INV_Tana_Local_Text, lstBoxIncremantal
    'イベント再開する
    StopEvents = False
End Sub
'TextBoxChange
'tana_System
Private Sub txtBox_F_INV_Tana_System_Text_Change()
    'tana_SystemがChangeしたら表示用のラベルのCaptionに同じ項目を設定してやる
    lblSystemTextView.Width = txtBox_F_INV_Tana_System_Text.Width
    lblSystemTextView.Caption = txtBox_F_INV_Tana_System_Text.Text
End Sub
Private Sub txtBox_F_INV_Tana_Local_Text_Change()
    On Error GoTo ErrorCatch
    If StopEvents Then
        'イベント停止状態なら抜ける
        Exit Sub
    End If
    'イベント停止する
    'UCASE掛ける
    txtBox_F_INV_Tana_Local_Text.Text = UCase(txtBox_F_INV_Tana_Local_Text.Text)
    'Incremental Textbox Change
    clsIncrementalNewLocation.Incremental_TextBox_Change True
    'LocalTextからスペースを除去し、SystemTextを得る
    Dim strSystemText As String
    'RegExpオブジェクトを定義
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    'パターンとして空白文字を設定
    '\s [\t\r\n\v\f]と等価
    objRegExp.Pattern = "\s"
    '文字列全体を検索
    objRegExp.Global = True
    strSystemText = objRegExp.REPLACE(txtBox_F_INV_Tana_Local_Text.Text, "")
    '一旦現在のRSのFilterを退避
    Dim varOldFilter As Variant
    varOldFilter = clsADONewLocation.RS.Filter
    'RSのFilterに取得したSystemTextをセットする
    clsADONewLocation.RS.Filter = dicObjToFieldName(txtBox_F_INV_Tana_System_Text.Name) & " = '" & strSystemText & "'"
    'Filter掛けた後RecordCountが1未満だったら新規レコード
    If clsADONewLocation.RS.RecordCount < 1 Then
        'tana_systemとしてスペースを空文字で置換したものをセット
        txtBox_F_INV_Tana_System_Text.Text = strSystemText
        '新規登録ボタンのEnableをTrueに
        btnAdNewLocation.Enabled = True
    Else
        '既存の棚番が存在する場合
        txtBox_F_INV_Tana_System_Text.Text = "(同名の棚番が存在します)"
        '新規登録ボタンのEnableをFalseに
        btnAdNewLocation.Enabled = False
    End If
    'フィルタを戻す
    clsADONewLocation.RS.Filter = varOldFilter
    'イベント再開する
    StopEvents = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "txtBox_F_INV_Tana_Local_Text_Change code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set objRegExp = Nothing
    Exit Sub
End Sub
'Click
'新規棚登録
Private Sub btnAdNewLocation_Click()
    AddNewLocation
End Sub
'''------------------------------------------------------------------------------------------------------
'プロシージャ
'''フォームコンストラクタ
Private Sub ConstRuctor()
    If clsADONewLocation Is Nothing Then
        Set clsADONewLocation = CreateclsADOHandleInstance
        '最初に明示的にDBPathをデフォルトに
        clsADONewLocation.SetDBPathandFilenameDefault
    End If
    If clsADONewLocation.RS Is Nothing Then
        Set clsADONewLocation.RS = New ADODB.Recordset
    End If
    If clsEnumNewLocation Is Nothing Then
        Set clsEnumNewLocation = CreateclsEnum
    End If
    If clsSqlBCNewLocation Is Nothing Then
        Set clsSqlBCNewLocation = CreateclsSQLStringBuilder
    End If
    If conADONewLocation Is Nothing Then
        Set conADONewLocation = New ADODB.Connection
    End If
    'dicObjeToFieldName と clsIncrementalは別プロシージャへ
    'dicObjToFieldNameの設定
    setdicObjToFieldName
    'デフォルトデータ取得
    setDefaultDatatoRS
    If clsIncrementalNewLocation Is Nothing Then
        Set clsIncrementalNewLocation = CreateclsIncrementalSerch
        'clsIncrementalのコンストラクタ
        clsIncrementalNewLocation.ConstRuctor frmRegistNewLocation, dicObjToFieldName, clsADONewLocation, clsEnumNewLocation, clsSqlBCNewLocation
    End If
End Sub
'デストラクタ
Private Sub DestRuctor()
    'クラスRS
    If Not clsADONewLocation.RS Is Nothing Then
        If clsADONewLocation.RS.State And ObjectStateEnum.adStateOpen Then
            '接続していたら切断してやる
            clsADONewLocation.RS.Close
        End If
        Set clsADONewLocation.RS = Nothing
    End If
    'Connection
    If Not conADONewLocation Is Nothing Then
        If conADONewLocation.State And ObjectStateEnum.adStateOpen Then
            conADONewLocation.Close
        End If
        Set conADONewLocation = Nothing
    End If
    If Not clsADONewLocation Is Nothing Then
        clsADONewLocation.CloseClassConnection
        Set clsADONewLocation = Nothing
    End If
    Unload frmRegistNewLocation
End Sub
'''dicObjToFieldの設定
Private Sub setdicObjToFieldName()
    If dicObjToFieldName Is Nothing Then
        Set dicObjToFieldName = New Dictionary
    End If
    '最初に全消去
    dicObjToFieldName.RemoveAll
    'コントロールとフィールド名対応付け
    dicObjToFieldName.Add lbl_F_INV_Tana_ID.Name, clsEnumNewLocation.INVMasterTana(F_INV_TANA_ID_IMT)
    dicObjToFieldName.Add txtBox_F_INV_Tana_System_Text.Name, clsEnumNewLocation.INVMasterTana(F_INV_Tana_System_Text_IMT)
    dicObjToFieldName.Add txtBox_F_INV_Tana_Local_Text.Name, clsEnumNewLocation.INVMasterTana(F_INV_Tana_Local_Text_IMT)
End Sub
'''DBよりデータ取得し、RSにセットする
Private Sub setDefaultDatatoRS()
    On Error GoTo ErrorCatch
    '置換用dic定義、設定
    Dim dicReplaceDefault As Dictionary
    Set dicReplaceDefault = New Dictionary
    dicReplaceDefault.RemoveAll
    dicReplaceDefault.Add 0, clsEnumNewLocation.INVMasterTana(F_INV_TANA_ID_IMT)
    dicReplaceDefault.Add 1, clsEnumNewLocation.INVMasterTana(F_INV_Tana_System_Text_IMT)
    dicReplaceDefault.Add 2, clsEnumNewLocation.INVMasterTana(F_INV_Tana_Local_Text_IMT)
    dicReplaceDefault.Add 3, INV_CONST.T_INV_M_Tana
    dicReplaceDefault.Add 4, PublicConst.INPUT_DATE
    'イベント停止する
    StopEvents = True
    'RSが未初期化の時は新たに設定してやる
    If clsADONewLocation.RS Is Nothing Then
        Set clsADONewLocation.RS = New ADODB.Recordset
    End If
    'RSかクラス変数のConnetcionオブジェクトが接続済みだったら一旦切断する
    If clsADONewLocation.RS.State And ObjectStateEnum.adStateOpen Then
        clsADONewLocation.RS.Close
    End If
    If conADONewLocation.State And ObjectStateEnum.adStateOpen Then
        conADONewLocation.Close
    End If
    'Connectionの設定
    conADONewLocation.ConnectionString = clsADONewLocation.CreateConnectionString(clsADONewLocation.DBPath, clsADONewLocation.DBFileName)
    conADONewLocation.CursorLocation = adUseClient
    conADONewLocation.Mode = adModeRead Or adModeShareDenyNone
    '接続オープン
    conADONewLocation.Open
    'RAのプロパティ設定
    clsADONewLocation.RS.LockType = adLockBatchOptimistic
    clsADONewLocation.RS.CursorType = adOpenStatic
    'Replace実行し、RSのSourceにSQL設定
    clsADONewLocation.RS.Source = clsSqlBCNewLocation.ReplaceParm(SQL_DEFAULT_NEW_LOCATION_DATA, dicReplaceDefault)
    'RSのConnectionに共有Connectionを設定
    Set clsADONewLocation.RS.ActiveConnection = conADONewLocation
    'rs Open
    clsADONewLocation.RS.Open , , , , CommandTypeEnum.adCmdText
    DebugMsgWithTime "setDefaultDatatoRS Default Data Count: " & clsADONewLocation.RS.RecordCount
    'イベント再開する
    StopEvents = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "getDefaultData code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set dicReplaceDefault = Nothing
    Exit Sub
End Sub
'新規棚番を登録する
Private Sub AddNewLocation()
    On Error GoTo ErrorCatch
    'イベント停止する
    StopEvents = True
    '現在のフィルターを退避
    Dim varOldFilter As Variant
    varOldFilter = clsADONewLocation.RS.Filter
    'SystemTextでFilterを設定
    clsADONewLocation.RS.Filter = dicObjToFieldName(txtBox_F_INV_Tana_System_Text.Name) & " = '" & txtBox_F_INV_Tana_System_Text.Text & "'"
    If clsADONewLocation.RS.RecordCount >= 1 Then
        'RecordCountが1以上の場合は既存のデータがあるので処理を中断
        MsgBox "既に同名の棚番が存在します。処理を中断します"
        'フィルタを戻す
        clsADONewLocation.RS.Filter = varOldFilter
        'TanaLocalにSetFocus
        txtBox_F_INV_Tana_Local_Text.SetFocus
        GoTo CloseAndExit
        Exit Sub
    End If
    'RSに新規レコードを追加
    clsADONewLocation.RS.AddNew
    '新規レコードに値をセットしていく
    Dim varKeydicObjt As Variant
    'dicObjToFieldをループ
    '空データ削除
    If dicObjToFieldName.Exists(Empty) Then
        dicObjToFieldName.Remove (Empty)
    End If
    For Each varKeydicObjt In dicObjToFieldName
        'オブジェクトの種類で処理を分岐
        Select Case True
        Case TypeName(frmRegistNewLocation.Controls(varKeydicObjt)) = "TextBox"
            'テキストボックスの場合
            '当面テキストボックスのみ値をセットする
            clsADONewLocation.RS.Fields(dicObjToFieldName(varKeydicObjt)).Value = frmRegistNewLocation.Controls(varKeydicObjt).Text
        End Select
    Next varKeydicObjt
    'InputDate入力
    clsADONewLocation.RS.Fields(PublicConst.INPUT_DATE).Value = GetLocalTimeWithMilliSec
    'ローカルのRSを確定させる
    clsADONewLocation.RS.Update
    '一旦フィルタ解除
    clsADONewLocation.RS.Filter = adFilterNone
    'PendingRecordsで変更のあるレコードのみでフィルタ
    clsADONewLocation.RS.Filter = adFilterPendingRecords
    'PendingRecordsのみでUpdateBatchをして、RSの結果をDBに反映させる
    Select Case True
    Case clsADONewLocation.RS.Status And (ADODB.RecordStatusEnum.adRecModified Or ADODB.RecordStatusEnum.adRecNew)
        'フィルタかけた後変更点あり、もしくは新規レコードだった場合
        'フィルタ条件に一致したもののみUpdate
        clsADONewLocation.RS.UpdateBatch adAffectGroup
        If clsADONewLocation.RS.Status And ADODB.RecordStatusEnum.adRecUnmodified Then
            MsgBox "正常に更新されました"
            '現在のレコードの値をフォームに反映させる
            GetValuFromRS
            '連続入力に対応するため、TanaLocalにSetFocus
            txtBox_F_INV_Tana_Local_Text.SetFocus
        Else
            DebugMsgWithTime "AddNewLocation : fail update batch Location_Local: " & txtBox_F_INV_Tana_Local_Text.Text
            MsgBox "正常に更新されなかった可能性があります 棚番: " & txtBox_F_INV_Tana_Local_Text.Text
        End If
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "AddNewLocation code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
'''RSより値を取得する
Private Sub GetValuFromRS()
    On Error GoTo CloseAndExit
    'イベント停止する
    StopEvents = True
    Dim varKeyDicObjtoField As Variant
    'cidObjToFieldの全要素をループ
    For Each varKeyDicObjtoField In dicObjToFieldName
        'コントロールの種類により処理を分岐
        Select Case True
        Case TypeName(frmRegistNewLocation.Controls(varKeyDicObjtoField)) = "Label"
            'ラベルだった場合
            frmRegistNewLocation.Controls(varKeyDicObjtoField).Caption = _
            clsADONewLocation.RS.Fields(dicObjToFieldName(varKeyDicObjtoField)).Value
        Case TypeName(frmRegistNewLocation.Controls(varKeyDicObjtoField)) = "TextBox"
            'テキストボックスだった場合
            frmRegistNewLocation.Controls(varKeyDicObjtoField).Text = _
            clsADONewLocation.RS.Fields(dicObjToFieldName(varKeyDicObjtoField)).Value
        End Select
    Next varKeyDicObjtoField
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "GetValuFromRS code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開する
    StopEvents = False
    Exit Sub
End Sub
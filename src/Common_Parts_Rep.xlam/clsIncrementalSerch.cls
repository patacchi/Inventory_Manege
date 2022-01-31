VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIncrementalSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''フォームの（主に）テキストボックスでインクリメンタルサーチを実装するためのクラス
Option Explicit
'メンバ変数
Private boolStopEvents As Boolean
Private varSourceArray As Variant                       'ListBoxに割り当てる大元の配列（フィルタ前、実際はこれにフィルタをかけてListBox.Listに割り当てる）
Private dicReplaceSELECTParam As Dictionary             'SQL SELECT用のReplaseParam
Private txtBoxTarget As MSForms.TextBox                 'ターゲットとなるテキストボックスの参照
Private lstBox_Incremental As MSForms.ListBox           'インクリメンタルサーチ用のリストボックスの参照
Private frmParent As UserForm                           'Parentとなるフォームそのもの
Private dicObjectToFieldName As Dictionary              'オブジェクト名をkey、テーブルエイリアス名付きのフィールド名をvalueとするDictionary
Private clsADOIncremental As clsADOHandle               'AdoHandleのインスタンス参照
Private clsEnumIncremental As clsEnum
Private clsSQLBIncremental As clsSQLStringBuilder
'メンバ定数
Private Const LISTBOX_FONT_SIZE As Single = 9           'リストボックスのフォントサイズ
Private Const LISTBOX_BASE_HEIGHT As Single = 10.5      'リストボックスの1行当たりの高さ
'プロパティ
'StopEvents
Property Get StopEvent() As Boolean
    StopEvent = boolStopEvents
    Exit Property
End Property
Property Let StopEvent(argBoolStopEvent As Boolean)
    boolStopEvents = argBoolStopEvent
    Exit Property
End Property
'txtBox
Property Get txtBoxRef() As MSForms.TextBox
    Set txtBoxRef = txtBoxTarget
End Property
Property Set txtBoxRef(ByRef argTxtBox As MSForms.TextBox)
    Set txtBoxTarget = argTxtBox
End Property
'lstBox
Property Get LstBoxRef() As MSForms.ListBox
    Set LstBoxRef = lstBox_Incremental
End Property
Property Set LstBoxRef(ByRef arglstBox As MSForms.ListBox)
    Set lstBox_Incremental = arglstBox
End Property
'me.frmref
Property Get frmRef() As UserForm
    Set frmRef = frmParent
End Property
Property Set frmRef(ByRef argFrm As UserForm)
    Set frmParent = argFrm
End Property
'me.dicobjectfield
Property Get dicObjectField() As Dictionary
    Set dicObjectField = dicObjectToFieldName
End Property
Property Set dicObjectField(ByRef argDicObject As Dictionary)
    Set dicObjectToFieldName = argDicObject
End Property
'clsADo
Property Get clsADOhandleRef() As clsADOHandle
    Set clsADOhandleRef = clsADOIncremental
End Property
Property Set clsADOhandleRef(ByRef argclsADO As clsADOHandle)
    Set clsADOIncremental = argclsADO
End Property
'clsEnum
Property Get clsEnumRef() As clsEnum
    Set clsEnumRef = clsEnumIncremental
End Property
Property Set clsEnumRef(ByRef argclsEnum As clsEnum)
    Set clsEnumIncremental = argclsEnum
End Property
'clsSqlBC
Property Get clsSQLBRef() As clsSQLStringBuilder
    Set clsSQLBRef = clsSQLBIncremental
End Property
Property Set clsSQLBRef(ByRef argclsSQLB As clsSQLStringBuilder)
    Set clsSQLBIncremental = argclsSQLB
End Property
'--------------------------------------------------------------------------------------------------------------
'メソッド
'''インクリメンタルサーチでテキストボックスのChangeイベントの共通イベントハンドラ
'''Changeイベント
'''文字が入力されたら、リストの絞り込みを行い、結果が1件のみになったらリストを消去する
'''args
Public Sub Incremental_TextBox_Change()
    On Error GoTo ErrorCatch
    'StopEventがセットされていたら何もしないで抜ける
    If Me.StopEvent Then
        'StopEvent
        Exit Sub
    End If
    '各引数のチェック
    Dim IsInitialize As Boolean
    IsInitialize = CheckInctance
    If Not IsInitialize Then
        'どれかのインスタンス参照変数がNothingだった
        GoTo CloseAndExit
        Exit Sub
    End If
    If Not Me.dicObjectField.Exists(Me.txtBoxRef.Name) Then
        'フィールド名Dicにテキストボックスの名前が見つからなかったら終了
        DebugMsgWithTime "Incremental_Textbox_Change: object name not found in Dictionary.  ObjectName: " & Me.txtBoxRef.Name
        GoTo CloseAndExit
        Exit Sub
    End If
    If Not IsRedim(varSourceArray) Then
        'まだ配列が初期化されてなかった場合
        GoTo CloseAndExit
    End If
    'イベント受付を停止する
    Me.StopEvent = True
    'リストボックスの選択状態を解除する
    Me.LstBoxRef.ListIndex = -1
    'テキストボックスに入力されている文字列でフィルタかけてみる
    Dim varArrayFilterd As Variant
    varArrayFilterd = Filter(varSourceArray, Me.txtBoxRef.Text, True, vbTextCompare)
    'filter関数は見つからないと 0 to -1 の配列を返すらしいので、Ubound < Lboundになるらしい
    If Not IsRedim(varArrayFilterd) Or UBound(varArrayFilterd) < LBound(varArrayFilterd) Then
        '見つからなかった場合？
        DebugMsgWithTime "Incremental_TextBox_Change: No result found"
        'リストボックスをクリアしてやり、VisibleをFalseにする
        Me.LstBoxRef.Clear
        Me.LstBoxRef.Visible = False
        GoTo CloseAndExit
        Exit Sub
    End If
    If UBound(varArrayFilterd) - LBound(varArrayFilterd) = 0 Then
        'フィルタ掛けた後の候補が残り1個になった場合
'        'リストボックスを非表示にして、リストボックスのIndex 0を選択し、Clickイベントを発火させる
''        'ここで非表示にしてしまうと検索文字列を削除した時にリスト復活しない
''        Me.LstBoxRef.Visible = False
'        Me.LstBoxRef.List = varArrayFilterd
'        Me.LstBoxRef.ListIndex = 0
'        'まずは表示調整をする
''        SetListBoxStyle varArrayFilterd
        '確定なので、リストボックスは表示しない→どうしよう?
        '候補1つになった時点で確定すると、次に連続入力したい時に不便
        'テキストボックスの値をリストボックスの値にセットしてやる
'        me.txtBoxRef.Text = me.lstboxref.List(0)
        'リストボックスのClickイベントを発生させるために選択をしてやる
'        me.lstboxref.ListIndex = 0
'        me.lstboxref.Visible = False
'        GoTo CloseAndExit
    End If
    'リストボックスを表示する
    SetListBoxStyle varArrayFilterd
    'イベント受付を再開する
    Me.StopEvent = False
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "Incremental_TextBox_Change code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Me.StopEvent = False
    Exit Sub
End Sub
'''インクリメンタルサーチ、テキストボックスイベントハンドラ
'''Enter
'''一番最初にこのイベントから始まるはずなので、クラスのコンストラクタ代わり
'''フォーカスが入った時点で、リストの初期化を行う
'''args
'''me.txtBoxRef                 ターゲットとなるテキストボックスの参照
'''me.lstboxref           インクリメンタルサーチ用のリストボックスの参照
'''me.frmref                    Parentのフォームの参照
'''me.dicobjectfield         オブジェクト名をkey、テーブルエイリアス名付きのフィールド名をvalueとするDictionary
'''me.clsadohandleref            AdoHandleのインスタンス参照
'''me.clsenumref
'''clsSQLStringBuilder
Public Sub Incremental_TextBox_Enter(ByRef txtBoxargTarget As MSForms.TextBox, ByRef lstBoxarg_Incremental As MSForms.ListBox, ByRef argFrmParent As UserForm, dicargObjectToFieldName As Dictionary, _
clsADOargIncremental As clsADOHandle, clsEnumargIncremental As clsEnum, clsSQLBargIncremental As clsSQLStringBuilder)
    On Error GoTo ErrorCatch
    '各引数のチェック、Nothingじゃなかったらクラスのプロパティにセットする
    If txtBoxargTarget Is Nothing Or lstBoxarg_Incremental Is Nothing Or dicargObjectToFieldName Is Nothing Or clsADOargIncremental Is Nothing _
     Or clsEnumargIncremental Is Nothing Or clsSQLBargIncremental Is Nothing Then
        'どれかのインスタンス参照がNothingだった
        DebugMsgWithTime "Incremental_Textbox_Enter: some instance Nothing"
        GoTo CloseAndExit
        Exit Sub
    Else
        'それぞれプロパティにセットしてやる
        Set Me.txtBoxRef = txtBoxargTarget
        Set Me.LstBoxRef = lstBoxarg_Incremental
        Set Me.frmRef = argFrmParent
        Set Me.dicObjectField = dicargObjectToFieldName
        Set Me.clsADOhandleRef = clsADOargIncremental
        Set Me.clsEnumRef = clsEnumargIncremental
        Set Me.clsSQLBRef = clsSQLBargIncremental
    End If
    If Not Me.dicObjectField.Exists(Me.txtBoxRef.Name) Then
        'フィールド名Dicにテキストボックスの名前が見つからなかったら終了
        DebugMsgWithTime "Incremental_Textbox_Enter: object name not found in Dictionary.  ObjectName: " & Me.txtBoxRef.Name
        GoTo CloseAndExit
        Exit Sub
    End If
    'リストの更新が終わるまでイベント停止する
    Me.StopEvent = True
    '一旦リストを非表示にし、選択を解除してやる
    Me.LstBoxRef.Visible = False
    Me.LstBoxRef.ListIndex = -1
    Me.LstBoxRef.Font.Size = LISTBOX_FONT_SIZE
    'ReplaceDicを設定する
    '今回はWhere条件はなし→X
    'Nullを排除しないとめんどうになる
    Dim strWhereCondition As String
    Dim strarrWhere(2) As String
    strarrWhere(0) = "AND NOT "
    strarrWhere(1) = Me.dicObjectField(Me.txtBoxRef.Name)
    strarrWhere(2) = " IS NULL"
    strWhereCondition = Join(strarrWhere, "")
    'ReplaceDicの設定を行う
    SetReplaceParmDic Me.dicObjectField(Me.txtBoxRef.Name), strWhereCondition
    'SQL組み立て
    Me.clsADOhandleRef.SQL = Me.clsSQLBRef.ReplaceParm(INV_CONST.SQL_INV_JOIN_TANA_PARTS, dicReplaceSELECTParam)
    Dim isCollect As Boolean
    'SQL実行
    isCollect = Me.clsADOhandleRef.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        DebugMsgWithTime "Incremental_TextBox_Enter: Do Get List sql fail..."
        MsgBox "Incremantal_TextBox_Enterでリスト取得SQL実行時にエラーが発生しました"
        GoTo CloseAndExit
    End If
    'タイトル無しの配列をメンバ変数に割り当てる
    'Filter関数では1次元配列しか引数にとれないので、メンバ変数に格納する前に1次元配列に変換してやる
    '一旦2次元配列として受け取る
    Dim vararr2Dimention As Variant
    vararr2Dimention = Me.clsADOhandleRef.RS_Array(NoTitle:=True)
    If Not IsRedim(vararr2Dimention) Then
        '初期化されてない配列が帰ってきたら
        MsgBox "データが見つかりませんでした"
        GoTo CloseAndExit
        Exit Sub
    End If
    '1次元配列に変換し、メンバ変数にセット
    varSourceArray = SQLResutArrayto1Dimmention(vararr2Dimention)
    '候補が一つの場合は直接テキストボックスに値を設定してやり、リストボックス自体は非表示のままにする
    If UBound(varSourceArray) - LBound(varSourceArray) = 0 Then
        Me.LstBoxRef.Visible = False
        Me.LstBoxRef.List = varSourceArray
        '1行目を選択し、本体のフォームでSelectイベントを走らせる
        'その後の処理は各フォームで行うものとする
        Me.LstBoxRef.ListIndex = 0
        'リストボックスを非表示にする
        Me.LstBoxRef.Height = 0
        Me.LstBoxRef.Visible = False
'        Me.txtBoxRef.Text = varSourceArray(0)
        GoTo CloseAndExit
        Exit Sub
    End If
    'テキストボックスに途中まで入力してまた戻ってきてもここに来るので、テキストボックスに文字が入っているかチェックし、あったらChangeハンドラに移動する
    If Me.txtBoxRef.TextLength >= 1 Then
        'イベント受付再開
        Me.StopEvent = False
        'Changeイベントハンドラ起動
        Incremental_TextBox_Change
        GoTo CloseAndExit
        Exit Sub
    End If
    'Enterの段階で既にテキストボックスに文字が入力されていたら
    'ここまでで大元のリストは取得できているので、一旦リスト表示できるように設定してやる
    SetListBoxStyle varSourceArray
    '絞り込みはChangeイベントにて行う
    'イベント受付再開
    Me.StopEvent = False
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "Incremental_TextBox_Enter Code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Me.StopEvent = False
    Exit Sub
End Sub
'''AdoHandleより返された2次元配列を1次元配列に置き換えて返す
Private Function SQLResutArrayto1Dimmention(varargArr2Demmention As Variant) As Variant
    On Error GoTo ErrorCatch
    If Not IsRedim(varargArr2Demmention) Or getArryDimmensions(varargArr2Demmention) <> 2 Then
        DebugMsgWithTime "arg array not AdoHandle Return array"
        GoTo CloseAndExit
    End If
    If UBound(varargArr2Demmention, 2) - LBound(varargArr2Demmention, 2) <> 0 Then
        '2次元目の要素数が1個以外だった
        DebugMsgWithTime "SWLResultArrayto1Dimmention: Dimmention 2 have not 1 elements"
        GoTo CloseAndExit
    End If
    '配列の要素数分ループ
    Dim vararr1Dimmention As Variant
    '配列のサイズを初期化
    ReDim vararr1Dimmention(UBound(varargArr2Demmention, 1))
    Dim longArrayRowCounter As Long
    For longArrayRowCounter = LBound(varargArr2Demmention, 1) To UBound(varargArr2Demmention, 1)
        vararr1Dimmention(longArrayRowCounter) = varargArr2Demmention(longArrayRowCounter, 0)
    Next longArrayRowCounter
    '結果を返す
    SQLResutArrayto1Dimmention = vararr1Dimmention
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SQLResultArrayto1Dimention code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
'''リストボックス(インクリメンタルリスト) Enterイベント
'''Enterした時の処理。残り1個だったらくりこ確定しちゃう位？
'''
Public Sub Incremental_LstBox_Enter()
    On Error GoTo ErrorCatch
    'インスタンス確認
    Dim isInitialized As Boolean
    isInitialized = CheckInctance
    If Not isInitialized Then
        '初期化が未完了？
        DebugMsgWithTime "Incremental_LstBox_Enter : some instance Nothing..."
        GoTo CloseAndExit
        Exit Sub
    End If
    If Me.LstBoxRef.ListCount = 1 Then
        'Enterした時点でリストの残数が1だった
        '選択してしまうのです
        '明示的にイベント再開
        Me.StopEvent = False
        Me.LstBoxRef.ListIndex = 0
        'ここに来る時点でClickイベントの処理は完了しちゃってるので、テキストボックスへのセットとリストボックスを非表示にしちゃっても大丈夫
        '一旦イベント停止
        Me.StopEvent = True
        'テキストボックスにテキストセット
        Me.txtBoxRef.Text = Me.LstBoxRef.List(Me.LstBoxRef.ListIndex)
        'リストボックス自体を非表示に、ついでにリストボックスの選択状態も解除する（次に選択されたままになってるため）
        Me.LstBoxRef.ListIndex = -1
        Me.LstBoxRef.Clear
        Me.LstBoxRef.Height = 0
        'イベント再開
        Me.StopEvent = False
        GoTo CloseAndExit
        Exit Sub
    End If
ErrorCatch:
CloseAndExit:
End Sub
'''インクリメンタルリストクリック（選択確定）イベント Click
'''指定されたDic情報をを元にSQL実行する。実行した後の値の設定はそれぞれの呼び出し元で行う。
'''戻り値   Bool    作業完了したらTrue、それ以外はFalse
'''呼び出し元ではTrueを待って値をセットしていくのかな？
Public Function Incremental_LstBox_Click() As Boolean
    On Error GoTo ErrorCatch
    '引数チェックをする
    Dim isInitialized As Boolean
    isInitialized = CheckInctance
    If Not isInitialized Then
        DebugMsgWithTime "Incremental_LstBox_Click : some inctance not initialized"
        Incremental_LstBox_Click = False
        GoTo CloseAndExit
    End If
    Dim strResultText As String
    '一旦イベントを停止する
    Me.StopEvent = True
    '結果を取得
    strResultText = Me.LstBoxRef.List(Me.LstBoxRef.ListIndex)
    'WHERE条件を作成
    Dim strWhereCondition As String
    Dim strarrCondition(3) As String
    strarrCondition(0) = "AND "
    strarrCondition(1) = Me.dicObjectField(Me.txtBoxRef.Name)
    strarrCondition(2) = " = "
    strarrCondition(3) = Me.clsSQLBRef.addQuote(strResultText)
    strWhereCondition = Join(strarrCondition, "")
    'ReplaceDicを設定する
    '今回のSELECTはDicObjctToFieldにある全ての項目を引っ張ってこないとダメ
    SetReplaceParmDic GetSELECTfieldListFromDicObjctToFieldName, strWhereCondition
    'clsADOのプロパティにSQLをセット
    Me.clsADOhandleRef.SQL = Me.clsSQLBRef.ReplaceParm(INV_CONST.SQL_INV_JOIN_TANA_PARTS, dicReplaceSELECTParam)
    'SQL実行
    Dim isCollect As Boolean
    isCollect = Me.clsADOhandleRef.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        DebugMsgWithTime "Incremental_LstBox_Click : SQL executing fali...."
        Incremental_LstBox_Click = False
        GoTo CloseAndExit
    End If
    'リストの選択状態解除、テキストボックスへの結果テキストの反映等はKeyUPおよびMouseUpイベント内で行う
    Incremental_LstBox_Click = True
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "Incremantal_LstBox_Click code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    'イベント再開
    Me.StopEvent = False
    Exit Function
End Function
'''インクリメンタルリストのKeyUPイベントのイベントハンドラ
'''KeyUpイベントなので、Clickイベントの後で発生するため自身を非表示にするのが主目的
'''args
'''argKeyCode       押されたキーのキーコードが入る
'''argShift         シフトキーの押下状態?
Public Sub Incremental_LstBox_Key_UP(argKeyCode As MSForms.ReturnInteger, argShift As Integer)
    On Error GoTo ErrorCatch
    'メンバ変数チェック
    Dim isInitialized As Boolean
    isInitialized = CheckInctance
    If Not isInitialized Then
        DebugMsgWithTime "Incremental_LstBox_Key_UP : Instance check fail"
        GoTo CloseAndExit
        Exit Sub
    End If
    'イベント受付停止する
    Me.StopEvent = True
    'キーコードにより処理を分岐
    Select Case argKeyCode
    Case vbKeyEscape
        'ESCキー
        'リストを非表示にする
        Me.LstBoxRef.Visible = False
        GoTo CloseAndExit
    Case vbKeyReturn
    Case vbKeyUp
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "Incremental_LstBox_Key_UP code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''T_INV_M_Tana と T_M_INV_Partsを F_INV_Tana_IDで結合した汎用SQLのRepleceDictionaryを設定するメソッド
'''args
'''strargSelectField        SELECT の引数
'''strargWhereCondition     WHERE の引数
Private Sub SetReplaceParmDic(strargSELECTField As String, strargWHERECondition As String)
    If dicReplaceSELECTParam Is Nothing Then
        Set dicReplaceSELECTParam = New Dictionary
''T_INV_M_TanaとT_INV_M_Parts結合した汎用SELECT SQL、外部DBファイル参照は無いものとする
''{0}    (SELECT Field)
''{1}    T_INV_M_Parts
''{2}    TDBPrts
''{3}    T_INV_M_Tana
''{4}    TDBTana
''{5}    F_INV_Tana_ID
''{6}    (WHERE condition)無しの場合は空文字 "" でOK
''INV_CONST.SQL_INV_JOIN_TANA_PARTS
        '初回は全ての条件を新たに入れてやる
        '0番SELECT フィールド
        dicReplaceSELECTParam.Add 0, strargSELECTField
        dicReplaceSELECTParam.Add 1, INV_CONST.T_INV_M_Parts
        dicReplaceSELECTParam.Add 2, Me.clsEnumRef.SQL_INV_Alias(INVDB_Parts_Alias_sia)
        dicReplaceSELECTParam.Add 3, INV_CONST.T_INV_M_Tana
        dicReplaceSELECTParam.Add 4, Me.clsEnumRef.SQL_INV_Alias(INVDB_Tana_Alias_sia)
        dicReplaceSELECTParam.Add 5, Me.clsEnumRef.INVMasterTana(F_INV_TANA_ID_IMT)
        '6番、WHERE条件
        dicReplaceSELECTParam.Add 6, strargWHERECondition
    Else
        '既に初期化済みの場合は、SELECT条件とWHERE条件のみ変更になるので、そこだけ消去して再割り当て
        '0 SELECT
        dicReplaceSELECTParam.Remove (0)
        dicReplaceSELECTParam.Add 0, strargSELECTField
        '6 WHERE
        dicReplaceSELECTParam.Remove (6)
        dicReplaceSELECTParam.Add 6, strargWHERECondition
    End If
End Sub
'''プロパティのdicObjToFieldNameより全フィールド選択のSELECT 引数を作成する
'''戻り値   string      SELECTの後にくっつけて使用する
Private Function GetSELECTfieldListFromDicObjctToFieldName() As String
    On Error GoTo ErrorCatch
    '引数チェック
    Dim isInitialized As Boolean
    isInitialized = CheckInctance
    If Not isInitialized Then
        DebugMsgWithTime "GetSELECTfieldListFromDicObjctToFieldName : not initialized instance var..."
        GetSELECTfieldListFromDicObjctToFieldName = ""
        GoTo CloseAndExit
    End If
    'dicobj.count分の配列を定義
    Dim strarrSELECTfield() As String
    ReDim strarrSELECTfield(Me.dicObjectField.Count - 1)
    Dim varKeyDicObj As Variant
    Dim longArrRowCounter As Long
    longArrRowCounter = 0
    For Each varKeyDicObj In Me.dicObjectField
        '配列にフィールド名を格納していく
        'フィールド名を明示的に指定しないとSELECTした結果がエイリアス外されちゃう
        'ついでにドットは特殊文字扱いされるのでRelaceしなきゃだめ
        Dim strarrPlusAlias(2) As String
        strarrPlusAlias(0) = Me.dicObjectField(varKeyDicObj)
        strarrPlusAlias(1) = " AS "
        strarrPlusAlias(2) = REPLACE(strarrPlusAlias(0), ".", "_")
        strarrSELECTfield(longArrRowCounter) = Join(strarrPlusAlias, "")
        '配列の行カウンターをインクリメント
        longArrRowCounter = longArrRowCounter + 1
    Next varKeyDicObj
    '結果を返す
    GetSELECTfieldListFromDicObjctToFieldName = Join(strarrSELECTfield, ",")
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "GetSELECTfieldListFromDicObjctToFieldName code: " & err.Number & " Descriptoin: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
'''ListBoxに各種プロパティをセットし、表示する
'''args
'''varArrayList                 リストボックスに表示する配列
'''me.txtBoxRef                 ターゲットとなるテキストボックスの参照
'''me.lstboxref           インクリメンタルサーチ用のリストボックスの参照
'''me.frmref                    Parentのフォームの参照
Private Sub SetListBoxStyle(varArrayList As Variant)
    On Error GoTo ErrorCatch
    If Not IsRedim(varArrayList) Or getArryDimmensions(varArrayList) <> 1 Then
        '結果の配列じゃないっぽいのがきた
        DebugMsgWithTime "SetListBoxstyle: parhaps not list array "
        Exit Sub
    End If
    Me.LstBoxRef.List = varArrayList
    With Me.LstBoxRef
        .ListStyle = fmListStylePlain
        .BorderStyle = fmBorderStyleSingle
        .Font.Size = LISTBOX_FONT_SIZE
        'サイズ調整し、テキストボックスのすぐ下に表示するようにする
        .Top = Me.txtBoxRef.Top + Me.txtBoxRef.Height
        .Width = Me.txtBoxRef.Width
        .Left = Me.txtBoxRef.Left
        .Height = LISTBOX_BASE_HEIGHT * .ListCount
        '高さの最大値の設定が必要かも？
        'フォームの高さを超えていたらHeightを調整
        If .Height > Me.frmRef.InsideHeight Then
            .Height = Me.frmRef.InsideHeight - .Top
        End If
        .Visible = True
    End With
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "SetListStyle code: " & err.Number & " Description : " & err.Description
    Exit Sub
CloseAndExit:
End Sub
'''クラス内のインスタンス参照変数の初期化状態を調べる関数
'''一つでもNothingがあればFalseを返す、全部初期化されていたらTrue
Private Function CheckInctance() As Boolean
    If Me.txtBoxRef Is Nothing Or Me.LstBoxRef Is Nothing Or Me.dicObjectField Is Nothing Or Me.clsADOhandleRef Is Nothing _
    Or Me.clsEnumRef Is Nothing Or Me.clsSQLBRef Is Nothing Or Me.frmRef Is Nothing Then
        'どれかのインスタンス参照がNothingだった
        DebugMsgWithTime "CheckInstance : some instance Nothing"
        CheckInctance = False
        GoTo CloseAndExit
        Exit Function
    Else
        '全てNothingじゃなかった
        CheckInctance = True
        GoTo CloseAndExit
        Exit Function
    End If
CloseAndExit:
    Exit Function
End Function
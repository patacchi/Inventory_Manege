VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'SQL文字操作（主に結合）のためのクラス
'開始履歴、終了履歴を設定して、文字列連結は配列のJoinで行う方向
Private strSQLBuilder As String         '最後の結果のSQL文
Private longStringLength As Long        '文字列の長さ
Private dblTimer As Double              '処理時間計測用タイマーアーアーア
'Private longRecordCount As Long         '現在の処理中データの場所（バルクの時に使う）
Private Sub Class_Initialize()
    'クラス初期化動作
    strSQLBuilder = Empty
    longStringLength = 0
End Sub
Private Sub Class_Terminate()
    'クラス破棄時に実行
    strSQLBuilder = Empty
    longStringLength = 0
End Sub
'''Author Daiske_Oota 2021_10_18
'''Printfのような関数
'''.NETのStringBuilderのAppendFormatのオーバーロード4を使用し
'''{index number 0start}形式のプレースホルダに穴埋めできる（ほかにも色々書式設定しようと思えばできる）
'''------------------------------------------------------------------------------------------------------
'''Parms
'''strargOriginString       プレースホルダ混みのオリジナル文字列
'''dicargParm               キーにインデックス番号、Itemにプレースホルダに入れるべき文字列を入れる。プレースホルダより数が少ない場合は例外がスローされる
Public Function ReplaceParm(strargOriginString As String, dicargParm As Dictionary) As String
    If strargOriginString = "" Then
        DebugMsgWithTime "ReplaceParm Orrigin String is empty"
        ReplaceParm = ""
        Exit Function
    End If
    If dicargParm.Count <= 0 Then
        '置き換え後のオブジェクト数がゼロだったら抜ける
        DebugMsgWithTime "ReplaceParm Parm list empty"
        ReplaceParm = ""
        Exit Function
    End If
    On Error GoTo ErrorCatch
    '引数のDictionaryを元にプレースホルダに対応するオブジェクトの配列を作成する
    Dim arrParm() As Variant
    ReDim arrParm(dicargParm.Count - 1)
    Dim longParmCount As Long
    For longParmCount = 0 To dicargParm.Count - 1
        arrParm(longParmCount) = dicargParm(longParmCount)
    Next longParmCount
    Dim objBuilder As Object
    Set objBuilder = CreateObject("System.Text.StringBuilder")
    ' System.Text.StringBuilder AppendFormat_4(string format, Params System.Object[] args)
    ReplaceParm = objBuilder.AppendFormat_4(strargOriginString, arrParm).ToString
'    objBuilder.Clear
    Set objBuilder = Nothing
    Exit Function
ErrorCatch:
    Set objBuilder = Nothing
    DebugMsgWithTime "ReplaceParm code: " & Err.Number & vbCrLf & " Description: " & Err.Description
    Exit Function
End Function
'''------------------------------------------------------------------------------------------------------
Private Function Append(ByVal strargOrignString As String, ByRef varAryString As Variant, Optional strDelimiter As String = ",") As String
    Dim strarrBuf() As String
    Dim longElementCount As Long
    Dim longCounter As Long
    On Error GoTo ErrorCatch
    Select Case Len(strargOrignString)
    Case 0
        '元のStringが長さ0の時は、追加の分が最初
        longElementCount = UBound(varAryString) - LBound(varAryString)
    Case Else
        '元のが文字ありだった場合
        '引数に指定されている要素数 + オリジナルのプラス一つ
        longElementCount = UBound(varAryString) - LBound(varAryString) + 1
    End Select
    ReDim strarrBuf(longElementCount)
    Select Case Len(strargOrignString)
    Case 0
        '元の長さが0の場合は追加の配列と同じでいい・・・
        strarrBuf = varAryString
    Case Else
        'オリジナルが文字ありだった場合
        'まず先頭にオリジナルのStringをくっつけてやる
        longCounter = 0
        strarrBuf(longCounter) = strargOrignString
        '追加分の配列を順番にくっつける
        For longCounter = 1 To longElementCount
            strarrBuf(longCounter) = varAryString(longCounter - (1 - LBound(varAryString)))
        Next longCounter
    End Select
    '配列に格納し終わったらJoinでStringにする
    Append = Join(strarrBuf, strDelimiter)
    Exit Function
ErrorCatch:
    DebugMsgWithTime "Append Error code: " & Err.Number & "Description: " & Err.Description
End Function
Public Function addQuote(ByVal strargOriginalString As String, Optional ByVal strQuote As String = """") As String
    '与えられた文字列に、エスケープ処理をする
    Dim strBuf(2) As String
    On Error GoTo ErrorCatch
    strBuf(0) = strQuote
    strBuf(1) = REPLACE(strargOriginalString, """", """""")
    strBuf(2) = strQuote
    addQuote = Append("", strBuf, "")
    Exit Function
ErrorCatch:
    DebugMsgWithTime "addQuote Error code: " & Err.Number & "Description: " & Err.Description
End Function
'''INVPartMasterテーブルと対応する在庫情報シートのフィールド名から UPDATE の SET パラメータや WHERE パラメータを作成する
'''対応するフィールド名を = や <> で連結したStringを返す
'''各フィールド名の前には、プレフィックスとしてテーブル名の短縮名エイリアスがつく
'''戻り値 String  ex) "TDBPrts.F_INV_Tehai_Text = TTmp.手配コード,TDBPrts.F_INV_Stock_Amount = TTmp.在庫数量・・・・"
'''args
'''PartsPrefix          PartsMasterのフィールドの前に付くプレフィックス
'''ZaikoSHPrefix        シートフィールド名の前につくプレフィックス、普通はTTmpになる
'''FieldCondition       フィールド間をつなぐ文字列の列挙隊の数値
'''ConnectEnum          複数条件時の条件間をつなぐ文字列 OR とか AND とか ,とか
'''clsEnumCondition     コンストラクタ動作軽量化のためのインスタンス共有
Public Function ReturnINVPartsandZaikoSHFieldCondition(PartsPrefix As Enum_SQL_INV_Alias, ZaikoSHPrefix As Enum_SQL_INV_Alias, FieldCondition As Enum_SQL_F_Condition, _
ConnectEnum As Enum_SQL_F_Condition, ByRef clsEnumCondition As clsEnum) As String
    On Error GoTo ErrorCatch
    'クラス変数チェック
    If clsEnumCondition Is Nothing Then
        'Enumが初期化されてなかった
        DebugMsgWithTime "ReturnArrayINVPartsandZaikoSHFieldCondition: Warning!! Enum instance not initialized.will delay..."
        Set clsEnumCondition = CreateclsEnum
    End If
    'まずはノーマルのフィールド一覧を受け取る(TRIM適用済み）x ここでTrimつけるのは早すぎる
    Dim arrstrNormalFields() As String
    arrstrNormalFields = clsEnumCondition.GetZaikoFieldNameArraySHandDB
    Dim longArrayRowCounter As Long
    '次に各要素にテーブル名プレフィックスをつける
    Dim arrstrAppendTablePreFix() As String
    ReDim arrstrAppendTablePreFix(UBound(arrstrNormalFields), 1)
    '受け取った配列のサイズ分ループ
    For longArrayRowCounter = LBound(arrstrNormalFields) To UBound(arrstrNormalFields)
        '連結用String配列
        Dim arrstrPrefix(2) As String
        'DBPartsPrefix
        arrstrPrefix(0) = clsEnumCondition.SQL_INV_Alias(PartsPrefix)
        arrstrPrefix(1) = "."
        arrstrPrefix(2) = arrstrNormalFields(longArrayRowCounter, 0)
        arrstrAppendTablePreFix(longArrayRowCounter, 0) = Join(arrstrPrefix, "")
        'SHPrefix
        arrstrPrefix(0) = clsEnumCondition.SQL_INV_Alias(ZaikoSHPrefix)
        arrstrPrefix(1) = "."
        arrstrPrefix(2) = arrstrNormalFields(longArrayRowCounter, 1)
        '条件によってはここでTrimをつけるかどうか判断する
        If clsEnumCondition.IsNeedTrimZaikoSh(arrstrNormalFields(longArrayRowCounter, 1)) Then
            'Trimが必要
            'プレフィックス、サフィックスを付加したものを返す
            Dim arrstrTrim(2) As String
            arrstrTrim(0) = clsEnumCondition.SQLFieldCondition(Trim_Prefix_sfc)
            arrstrTrim(1) = Join(arrstrPrefix, "")
            arrstrTrim(2) = clsEnumCondition.SQLFieldCondition(Trim_Suffix_sfc)
            arrstrAppendTablePreFix(longArrayRowCounter, 1) = Join(arrstrTrim, "")
        Else
            'Trimは不要
            'そのまま結合し、結果を返す
            arrstrAppendTablePreFix(longArrayRowCounter, 1) = Join(arrstrPrefix, "")
        End If
    Next longArrayRowCounter
    '次に演算子を結合する
    '1次元に集約する
    Dim arrstrOperatorAdd() As String
    ReDim arrstrOperatorAdd(UBound(arrstrAppendTablePreFix))
    For longArrayRowCounter = LBound(arrstrAppendTablePreFix) To UBound(arrstrAppendTablePreFix)
        Dim arrstrOperator(2) As String
        arrstrOperator(0) = arrstrAppendTablePreFix(longArrayRowCounter, 0)
        arrstrOperator(1) = clsEnumCondition.SQLFieldCondition(FieldCondition)
        arrstrOperator(2) = arrstrAppendTablePreFix(longArrayRowCounter, 1)
        arrstrOperatorAdd(longArrayRowCounter) = Join(arrstrOperator, "")
    Next longArrayRowCounter
    '各条件を結合文字列で区切ってStringとして渡すが、要素が１個しかない時は不要なので処理を分岐させる
    If (UBound(arrstrOperatorAdd) - LBound(arrstrOperatorAdd)) >= 1 Then
        '要素数が2個以上ある時
        '要素を結合文字列で区切って渡す
        ReturnINVPartsandZaikoSHFieldCondition = Join(arrstrOperatorAdd, clsEnumCondition.SQLFieldCondition(ConnectEnum))
    Else
        '要素が1個だけだった場合
        '最初の要素をそのまま返す
        ReturnINVPartsandZaikoSHFieldCondition = arrstrOperatorAdd(LBound(arrstrOperatorAdd))
    End If
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "ReturnINVPartsAndZaikoSHFieldCondition code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
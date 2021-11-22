VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    strBuf(1) = Replace(strargOriginalString, """", """""")
    strBuf(2) = strQuote
    addQuote = Append("", strBuf, "")
    Exit Function
ErrorCatch:
    DebugMsgWithTime "addQuote Error code: " & Err.Number & "Description: " & Err.Description
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGetIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''IEのコンテンツをVBAで扱えるようにするクラスです
'''--------------------------------------------------------------------------------------------------------------
'''必須参照設定
'''Microsoft Internet Controles
'''Microsoft HTML Object Libraly
Private Const START_ROW_DEFAULT As Long = 6                                 'シート書き出しの際のデフォルトスタート行
Private Const START_COLUMN_DEFAULT As Long = 2                              'シート書き出しのデフォルトスタート列
'IE接続に関するメンバ変数
Private IEObject As InternetExplorer                                        'IEオブジェクト
Private strURL As String                                                    'IEで開くURLを保持するメンバ変数
Private isVisible As Boolean                                                'IE画面を表示するかどうか、Trueで表示
Private tableRow As HTMLTableRow                                            'テーブルの<tr>の内容、行情報
Private tableData As HTMLTableCell                                          'テーブルの<td>の内容、個別のデータ
'シート書き出しの際に使用するメンバ変数
Private longColumn As Long                                                  'シート書き出す際の列番号
Private longRow As Long                                                     'シート書き出す際の行番号
'プロパティ（アクセサ）
'URL
Property Get URL() As String
    URL = strURL
End Property
Property Let URL(strargURL As String)
    If Not strargURL = "" Then
        strURL = strargURL
    End If
End Property
'Column
Property Get shColumn() As Long
    '0以下の場合はデフォルト値をセットしてやる
    If longColumn <= 0 Then
        longColumn = START_COLUMN_DEFAULT
    End If
    shColumn = longColumn
End Property
Property Let shColumn(longargColumn As Long)
    '当面1以上の数のみ受付
    If longargColumn >= 1 Then
        longColumn = CLng(longargColumn)
    End If
End Property
'Row
Property Get shRow() As Long
    '0以下の場合はデフォルト値をセットしてやる
    If longRow <= 0 Then
        longRow = START_ROW_DEFAULT
    End If
    shRow = longRow
End Property
Property Let shRow(longargRow As Long)
    '当面1以上の数のみ受付
    If longargRow >= 1 Then
        longRow = CLng(longargRow)
    End If
End Property
'ieObject 値を設定するのに必要かも
Property Get IEInstance() As InternetExplorer
    Set IEInstance = IEObject
    Exit Property
End Property
'Visible
Property Get Visible() As Boolean
    Visible = isVisible
End Property
Property Let Visible(boolargVisible As Boolean)
    isVisible = boolargVisible
End Property
'コンストラクタ
'実際の処理内容はメソッドに分離する
Private Sub Class_Initialize()
    Call Initialize
End Sub
'デストラクタ
Private Sub Class_Terminate()
    Call Finalize
End Sub
'実際の処理内容はメソッドに分離する
'メソッド
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''初期化時に実行する処理
'''--------------------------------------------------------------------------------------------------------------
Private Sub Initialize()
    longColumn = START_COLUMN_DEFAULT
    longRow = START_ROW_DEFAULT
'    Set ieObject = New InternetExplorer
    '
    Set IEObject = New InternetExplorerMedium
    isVisible = False
End Sub
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''クラス破棄時に実行する処理
'''--------------------------------------------------------------------------------------------------------------
Private Sub Finalize()
    On Error GoTo ErrorCatch
    If Not IEObject Is Nothing Then
        IEObject.Quit
    End If
    Set IEObject = Nothing
    Exit Sub
ErrorCatch:
    Set IEObject = Nothing
    Exit Sub
End Sub
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''指定されたURLをブラウザで開く
'''--------------------------------------------------------------------------------------------------------------
Public Function OpenIEwithURL(Optional strargURL As String) As Boolean
    'URLが指定されていなかったら即抜ける
    If strargURL = "" Then
        strargURL = strURL
    End If
    If strargURL = "" Then
        MsgBox "URLが指定されていません（必須）"
        Debug.Print "OpenIEwithURL: URL is empty"
        OpenIEwithURL = False
        Exit Function
    End If
    'isVisibleプロパティの設定値を反映させる
    IEObject.Visible = isVisible
    IEObject.navigate strargURL
End Function
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''指定されたURLを元にHTMLDocumentオブジェクトを値に持つDictionaryを返す
'''キーとしてフレーム番号を取る
'''親ドキュメントは "top"をキーとして、格納する
'''--------------------------------------------------------------------------------------------------------------
'''Parms
'''strargURL        IEに渡すURLを指定
Public Function ReturnHTMLDocbyURL(Optional strargURL As String) As Dictionary
    Dim localDic As Dictionary
    Dim longElementCount As Long
    Dim intFrameCount As Integer
    Dim topDoc As HTMLDocument
    Dim frameDoc As HTMLDocument
    '引数にURLが指定されていなくて、プロパティで指定がある場合は、そちらを使ってやる
    If strargURL = "" And Not (strURL = "") Then
        strargURL = strURL
    End If
    If strargURL = "" Then
        MsgBox "URLが指定されていません（必須）"
        Debug.Print "OpenIEwithURL: URL is empty"
        Set ReturnHTMLDocbyURL = localDic
        GoTo CloseAndExit
    End If
    '指定のページを開く
    OpenIEwithURL (strargURL)
'    ieObject.navigate strargURL
    '親ページ読み込み完了まで待機する
    Application.StatusBar = "TOPページ読み込み完了待ち"
    Do Until IEObject.Busy = False And IEObject.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.StatusBar = "TOPページ読み込み完了。子ページ読み込みに入ります"
    Set topDoc = IEObject.document
    'HTMLDocumentの読み込み完了を待ってから、フレームの有無について確認し
    'フレームが存在する場合は、フレーム内の読み込み完了も待機する
    Call IsReadComp(topDoc)
    Application.StatusBar = "TOPページ準備完了"
    '2021_11_18 Daisuke_oota
    '認証チェック
    'ここで、指定されていたURLとieobjectのLocationが違っていた場合は、ログイン画面にリダイレクトされてると思うので
    '一旦VisibleをTrueにしてやって、ログイン処理をしてもらう
    If strargURL <> IEObject.LocationURL Then
        Debug.Print "URL不一致、リダイレクトされた可能性があります"
        '強制的にVisibleをTrueにする
        IEObject.Visible = True
        'ログイン処理完了したら元のページに戻るはずなので、ieobjectのLocationURLが指定されたものになるまで待機
        'サイトによっては無限ループの可能性有り・・・
        Do Until IEObject.LocationURL = strargURL And IEObject.Busy = False And IEObject.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        'Visibleの状態をプロパティで設定してる状態に戻してやる
        IEObject.Visible = Me.Visible
    End If
    If topDoc Is Nothing Then
        'もしこの時点でHTMLDocがNothingだったら何かあったと思うので、以後の処理は中止する
        Debug.Print "ReturnHTMLDocbyURL: topHTMLDoc is nothing"
        Set ReturnHTMLDocbyURL = localDic
    End If
    'この時点で親ドキュメントは読み込み完了しているので、Dictionaryに追加してやる
    Set localDic = New Dictionary
    localDic.Add "top", topDoc
    intFrameCount = 0
    If topDoc.frames.Length > 0 Then
        'フレームが存在する場合
        Application.StatusBar = "フレームが存在するのでフレームの読み込み完了待機に入ります。"
        For intFrameCount = 0 To topDoc.frames.Length - 1
            'フレームの存在する数分ループ
            On Error Resume Next
                Set frameDoc = topDoc.frames(intFrameCount).document
                'エラー発生してたらデバッグウィンドウに表示してやる
                If Err.Number <> 0 Then
                    Debug.Print "ReturnHTMLDocbyURL code: " & Err.Number & " Description: " & Err.Description & " FrameCount: " & intFrameCount
                End If
            On Error GoTo 0
            If Not frameDoc Is Nothing Then
                'フレームの中身があるみたい
                Application.StatusBar = intFrameCount & " 番目のフレーム読み込み完了待ち"
                'フレームの読み込み完了を待つ
                Call IsReadComp(frameDoc)
                'Dictionaryに追加してやる
                localDic.Add intFrameCount, frameDoc
                Set frameDoc = Nothing
            End If
        Next intFrameCount
        Application.StatusBar = "フレーム読み込完了"
        '結果を返してやる
        Set ReturnHTMLDocbyURL = localDic
        Application.StatusBar = "結果書き出し完了"
        GoTo CloseAndExit
    Else
        'フレームが存在しない場合
'        'topのみをDictionaryに入れて返す
'        localDic.Add intFrameCount, topDoc
        Set ReturnHTMLDocbyURL = localDic
        GoTo CloseAndExit
        Exit Function
    End If
CloseAndExit:
    Set localDic = Nothing
    Exit Function
End Function
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''HTMLドキュメントを引数として、読み込み完了まで待機し、完了したらTrueを返す
'''--------------------------------------------------------------------------------------------------------------
'''Parms
'''argHTMLDoc           読み込み完了を待つHTMLドキュメント
Private Function IsReadComp(argHTMLDoc As HTMLDocument) As Boolean
    Application.StatusBar = argHTMLDoc.Title & " のページ読み込み完了を待機しています "
    If Not argHTMLDoc Is Nothing Then
        'HTMLDocが中身ある時だけ待機ループする
        Do Until argHTMLDoc.readyState = "complete"
            DoEvents
        Loop
    End If
    Application.StatusBar = "読み込み完了"
    IsReadComp = True
    Exit Function
End Function
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''HTMLドキュメントを引数として
'''指定された親タグの中の(例：Table)子タグ(th)の中身(Text)を取得する
'''戻り値はString配列
'''--------------------------------------------------------------------------------------------------------------
'''Parms
Public Function getTextArrayByTagName(argHTMLDoc As HTMLDocument, argTopTagName As String, Optional argChildTagName As String = "") As String()
    '親タグの要素を取得する
    If (argHTMLDoc Is Nothing) Or (argTopTagName = "") Then
        Debug.Print "getTextArrayByTagName HTMLDocument or TopTagName is empty"
    End If
    Dim localarrString() As String
    On Error GoTo ErrorCatch
    Dim parentElement As Object
'    Dim parentElement As HTMLElementCollection
    Set parentElement = argHTMLDoc.getElementsByTagName(argTopTagName)
    If parentElement.Length = 0 Then
        '要素数が0の場合はここで抜ける
        getTextArrayByTagName = localarrString
        Set parentElement = Nothing
        Exit Function
    End If
    ReDim localarrString(parentElement.Length)
    Dim varElement As Variant
    Dim longRowCount As Long
    longRowCount = 0
    For Each varElement In parentElement
'        Debug.Print varElement.innerText
        'ここで要素のTextをString配列に追加していってやる予定
        localarrString(longRowCount) = varElement.innerText
        longRowCount = longRowCount + 1
    Next varElement
    '子タグの引数にデータが入っていれば更に子タグの要素も取得する
    If argChildTagName <> "" Then
        Dim childElement As HTMLElementCollection
    End If
    getTextArrayByTagName = localarrString
    Exit Function
ErrorCatch:
    Debug.Print "getTestArrayByTagNae code: " & Err.Number & " Description: " & Err.Description
    getTextArrayByTagName = localarrString
    Exit Function
End Function
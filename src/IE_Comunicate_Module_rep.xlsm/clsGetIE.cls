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
Private Const TOP_DOC_CHR As String = "d"                                   '結果格納用Dictionaryのトップフレームの文字列
'IE接続に関するメンバ変数
Private ieObject As InternetExplorer                                        'IEオブジェクト
Private strURL As String                                                    'IEで開くURLを保持するメンバ変数
Private isVisible As Boolean                                                'IE画面を表示するかどうか、Trueで表示
Private dicReturnHTMLDocs As Dictionary                                     '各フレーム毎にHTML Doumentを入れるkeyはフレーム番号(t001・・・),valueはHTML Document
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
Property Let URL(strargUrl As String)
    If Not strargUrl = "" Then
        strURL = strargUrl
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
'ieObjectを返す
Property Get IEInstance() As InternetExplorer
    Set IEInstance = ieObject
    Exit Property
End Property
'Visible
Property Get Visible() As Boolean
    Visible = isVisible
End Property
Property Let Visible(boolargVisible As Boolean)
    isVisible = boolargVisible
End Property
'dicHTMLDocs
'読み取り専用
Property Get ResultHTMLDoc(Optional ByRef argstrURL As String = "") As Dictionary
    '引数に与えられたURLを元にメンバ変数のdicReturnHTMLDocsを更新してみる
    If dicReturnHTMLDocs Is Nothing Then
        Set dicReturnHTMLDocs = New Dictionary
    End If
    SetHTMLDocbyURL (argstrURL)
    Set ResultHTMLDoc = dicReturnHTMLDocs
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
    Set dicReturnHTMLDocs = New Dictionary
    Set ieObject = New InternetExplorerMedium
    isVisible = False
End Sub
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''クラス破棄時に実行する処理
'''--------------------------------------------------------------------------------------------------------------
Private Sub Finalize()
    On Error GoTo ErrorCatch
    If Not ieObject Is Nothing Then
        ieObject.Quit
    End If
    Set ieObject = Nothing
    Exit Sub
ErrorCatch:
    Set ieObject = Nothing
    Exit Sub
End Sub
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''指定されたURLをブラウザで開く
'''--------------------------------------------------------------------------------------------------------------
Public Function OpenIEwithURL(Optional strargUrl As String) As Boolean
    'URLが指定されていなかったら即抜ける
    If strargUrl = "" Then
        strargUrl = strURL
    End If
    If strargUrl = "" Then
        MsgBox "URLが指定されていません（必須）"
        DebugMsgWithTime "OpenIEwithURL: URL is empty"
        OpenIEwithURL = False
        Exit Function
    End If
    'isVisibleプロパティの設定値を反映させる
    ieObject.Visible = isVisible
    ieObject.navigate strargUrl
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
Public Sub SetHTMLDocbyURL(Optional strargUrl As String)
    Dim longElementCount As Long
    Dim intFrameCount As Integer
    Dim topDoc As HTMLDocument
    Dim frameDoc As HTMLDocument
    '引数にURLが指定されていなくて、プロパティで指定がある場合は、そちらを使ってやる
    If strargUrl = "" And Not (strURL = "") Then
        strargUrl = strURL
    End If
    If strargUrl = "" Then
        MsgBox "URLが指定されていません（必須）"
        DebugMsgWithTime "OpenIEwithURL: URL is empty"
        GoTo CloseAndExit
    End If
    '指定のページを開く
    OpenIEwithURL (strargUrl)
'    ieObject.navigate strargURL
    '親ページ読み込み完了まで待機する
    Application.StatusBar = "TOPページ読み込み完了待ち"
    Do Until ieObject.Busy = False And ieObject.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.StatusBar = "TOPページ読み込み完了。子ページ読み込みに入ります"
    Set topDoc = ieObject.document
    'HTMLDocumentの読み込み完了を待ってから、フレームの有無について確認し
    'フレームが存在する場合は、フレーム内の読み込み完了も待機する
    Call IsReadComp(topDoc)
    Application.StatusBar = "TOPページ準備完了"
    '2021_11_18 Daisuke_oota
    '認証チェック
    'ここで、指定されていたURLとieobjectのLocationが違っていた場合は、ログイン画面にリダイレクトされてると思うので
    '一旦VisibleをTrueにしてやって、ログイン処理をしてもらう
    If strargUrl <> ieObject.LocationURL Then
        DebugMsgWithTime "URL不一致、リダイレクトされた可能性があります"
        '強制的にVisibleをTrueにする
        ieObject.Visible = True
        'ログイン処理完了したら元のページに戻るはずなので、ieobjectのLocationURLが指定されたものになるまで待機
        'サイトによっては無限ループの可能性有り・・・
        Do Until ieObject.LocationURL = strargUrl And ieObject.Busy = False And ieObject.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        'Visibleの状態をプロパティで設定してる状態に戻してやる
        ieObject.Visible = Me.Visible
    End If
    If topDoc Is Nothing Then
        'もしこの時点でHTMLDocがNothingだったら何かあったと思うので、以後の処理は中止する
        DebugMsgWithTime "SetHTMLDocbyURL: topHTMLDoc is nothing"
'        Set SetHTMLDocbyURL = localDic
    End If
    'この時点で親ドキュメントは読み込み完了しているので、Dictionaryに追加してやる
    AddResultdicHTML topDoc
CloseAndExit:
    Exit Sub
End Sub
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''HTMLドキュメントを引数として、クラスのメンバ変数のDictionaryにフレーム階層文字列とHTMLDocumentを追加してやる
'''子フレームが見つかった場合は再帰呼び出しを行う
'''--------------------------------------------------------------------------------------------------------------
'''Parms
'''atrHTMLDoc   その時点でのフレームのトップドキュメント
'''strTreeChr  フレームの階層文字列 トップはt（Const定義によ）その後は、framesのインデックス番号を階層ごとに追加してやる
Private Sub AddResultdicHTML(ByRef argHTMLDoc As HTMLDocument, Optional ByRef strTreeChr As String = "")
    If strTreeChr = "" Then
        'フレーム階層文字が空白だったらtopドキュメントのはずなのでconstのトップど階層文字列を設定してやる
        strTreeChr = TOP_DOC_CHR
    Else
        '階層文字に指定があった場合は、１階層下のトップなので、0を後ろに付加してやる
        strTreeChr = strTreeChr & TOP_DOC_CHR
    End If
    'この時点でトップドキュメントをDictionaryに追加してやる
    If Not argHTMLDoc Is Nothing Then
        dicReturnHTMLDocs.Add strTreeChr, argHTMLDoc
    End If
    If argHTMLDoc.frames.Length > 0 Then
        'フレームが存在する場合
        Application.StatusBar = "フレームが存在するのでフレームの読み込み完了待機に入ります。"
        'フレーム数分ループする
        Dim intFrameCount As Integer
        Dim frameDoc As HTMLDocument
        For intFrameCount = 0 To argHTMLDoc.frames.Length - 1
            'セキュリティの設定でドメインが違うと例外を吐くようなので、とりあえず無視する
            On Error Resume Next
            Set frameDoc = argHTMLDoc.frames(intFrameCount).document
            'エラー発生してたらデバッグ出力する
            If Err.Number <> 0 Then
                DebugMsgWithTime "AddResultdicHTML code: " & Err.Number & " Description: " & Err.Description
            End If
            On Error GoTo 0
            If Not frameDoc Is Nothing Then
                '中身がある場合は
                Application.StatusBar = intFrameCount & " 番目のフレームの読み込み完了待ち"
                'フレームの読み込み完了を待つ
                Call IsReadComp(frameDoc)
                'まずは階層文字に現在のフレームカウントを付加して、Dictionaryに追加してやる
                Dim strlocalFrameNum As String
                strlocalFrameNum = strTreeChr & CStr(intFrameCount)
                If dicReturnHTMLDocs.Exists(strlocalFrameNum) Then
                    '同じキーが存在してたら・・・んー・・？
                    DebugMsgWithTime "AddResultdicHTML: same FrameNumStr available"
                    'とりあえず処理中断
                    Stop
                End If
                dicReturnHTMLDocs.Add strlocalFrameNum, frameDoc
                'ここで現在のフレームに更に子フレームがあるか判断し
                'ある場合は再帰呼び出しを行う
                If frameDoc.frames.Length > 0 Then
                    '更にフレームがあったので現在のフレーム階層文字列を引数として再帰呼び出し
                    AddResultdicHTML frameDoc, strlocalFrameNum
                End If
            Else
                '中身がない場合は、このループはもう何もしない
            End If
            '次のループのために初期化してやる
            Set frameDoc = Nothing
        Next intFrameCount
        Application.StatusBar = "フレーム読み込完了"
    Else
        'フレームが存在しない場合
        '最初でトップドキュメントを追加してるので､ここでは何もしない
    End If
End Sub
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
'''指定された親タグの中の(例：Table)子タグ(th)の中身(OuterHTML)を取得する
'''戻り値はString配列
'''--------------------------------------------------------------------------------------------------------------
'''Parms
Public Function getTextArrayByTagName(argHTMLDoc As HTMLDocument, argTopTagName As String, Optional argChildTagName As String = "") As String()
    '親タグの要素を取得する
    If (argHTMLDoc Is Nothing) Or (argTopTagName = "") Then
        DebugMsgWithTime "getTextArrayByTagName HTMLDocument or TopTagName is empty"
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
'        DebugMsgWithTime varElement.innerText
        'ここで要素のTextをString配列に追加していってやる予定
        localarrString(longRowCount) = varElement.outerHTML
        longRowCount = longRowCount + 1
    Next varElement
    '子タグの引数にデータが入っていれば更に子タグの要素も取得する
    If argChildTagName <> "" Then
        Dim childElement As HTMLElementCollection
    End If
    getTextArrayByTagName = localarrString
    Exit Function
ErrorCatch:
    DebugMsgWithTime "getTestArrayByTagNae code: " & Err.Number & " Description: " & Err.Description
    getTextArrayByTagName = localarrString
    Exit Function
End Function
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''dicHTMLResut (key: TreeChr、階層文字、value:HTML Document)、タグ名を引数として、中のOuter HTMLをDictionaryで返す
'''戻り値はDictionary( key: TreChr, Value string配列)
'''--------------------------------------------------------------------------------------------------------------
'''Parms
'''dicHTMLResult        Key:TreeChr Value:HTML Document のDictionary
'''strTagName           検索するタグ名
Public Function GetHTMLdicBydicHTMLDocandTagName(ByRef argdicHTML As Dictionary, ByVal strTagName As String) As Dictionary
    If argdicHTML Is Nothing Or strTagName = "" Then
        '引数が空だったよ
        DebugMsgWithTime "GetHTMLstrngbyDicHTMLDocandTagName: dicHTML or TagName empty"
        Set GetHTMLdicBydicHTMLDocandTagName = New Dictionary
        Exit Function
    End If
    Dim currentDoc As HTMLDocument
    Dim keyDicHTML As Variant
    Dim dicResult As Dictionary
    Set dicResult = New Dictionary
    On Error GoTo ErrorCatch
    For Each keyDicHTML In argdicHTML
        Set currentDoc = argdicHTML(keyDicHTML)
        If Not currentDoc Is Nothing Then
            '中身ある時だけ続ける
            Dim arrstrElements() As String
            arrstrElements() = getTextArrayByTagName(currentDoc, strTagName)
            If IsRedim(arrstrElements) Then
                '初期化済みの配列が帰ってきたら、Dictionary の Valueにセットしてやる
                dicResult.Add keyDicHTML, arrstrElements
            End If
            '次のループのために配列を初期化する
            Erase arrstrElements
        End If
    Next keyDicHTML
    Set GetHTMLdicBydicHTMLDocandTagName = dicResult
    Set dicResult = Nothing
    Exit Function
ErrorCatch:
    DebugMsgWithTime "GetHTMLdicBydicHTMLDocandTagName: " & Err.Number & " Description: " & Err.Description
    Set GetHTMLdicBydicHTMLDocandTagName = dicResult
    Exit Function
End Function
Attribute VB_Name = "IE_Connect_Test"
Option Explicit
Private Const START_COLUMN As Long = 2
Private Const START_ROW As Long = 4
Public Function OpenUrlatIE(Optional ByVal strargUrl As String = "") As Boolean
'    Dim ieObject As InternetExplorerMedium
    Dim ieObject As InternetExplorer
    Dim returnHTML As HTMLDocument
'    Dim tableWhole As IHTMLElementCollection
    Dim tableRow As HTMLTableRow
    Dim tableHeader As HTMLTableCell
    Dim tableData As HTMLTableCell
    Dim longColumnCount As Long
    Dim longRowCount As Long
    Set ieObject = New InternetExplorer
    ieObject.Visible = False
    If strargUrl = "" Then
'        '沖縄県市町村一覧ダウンロード_SaveAsテスト
'        strargUrl = "https://saigai.gsi.go.jp/jusho/download/pref/47.html"
        'confirm強行突破テスト用
        strargUrl = "http://needtec.sakura.ne.jp/auto_demo/form1.html"
    End If
    ieObject.navigate strargUrl
    '読み込み完了まで待つ処理
    Do While ieObject.Busy = True And ieObject.readyState <> READYSTATE_COMPLETE
        Application.StatusBar = "トップ画面読み込み完了待ち"
        DoEvents
    Loop
    Application.StatusBar = ""
    'HTMLDobumentオブジェクトとして取得
    Set returnHTML = ieObject.document
    '読み込んだドキュメントの読み込み完了を待機
    Do While returnHTML.readyState <> "complete"
        Application.StatusBar = "Document読み込み完了待機中..."
        DoEvents
    Loop
    Application.StatusBar = "読み込み完了"
    ieObject.Visible = True
    Dim htmlelmName As HTMLDocument
    Set htmlelmName = returnHTML.getElementsByName("name").Item(, 0)
    htmlelmName.Value = "ぽにぷに"
    Dim htmlelmMail As HTMLDocument
    Set htmlelmMail = returnHTML.getElementsByName("mail").Item(, 0)
    htmlelmMail.Value = "puni@poni"
    'confirm偽造
    returnHTML.parentWindow.execScript "confirm = function(){return true;}"
    Dim htmlelmSubmitButton As Object
    Set htmlelmSubmitButton = returnHTML.getElementsByTagName("input")
    Dim objelm As Object
    For Each objelm In htmlelmSubmitButton
        If InStr(objelm.outerHTML, "登録する") >= 1 Then
            objelm.Click
        End If
    Next objelm
''-----------------------------------------------------------------------------------------------------------------------------
''NotificationSaveAs 使用例
'    糸満市のデータが軽いのでそのリンクを探す
'    https://saigai.gsi.go.jp/jusho/download/data/47210.zip <a href>
'    Dim htmlLiks As HTMLHtmlElement
'    HTMLDoc.Liks で aタグのhrefの一覧を取得できるそう
'    Dim fsoLink As FileSystemObject
'    For Each htmlLiks In returnHTML.Links
'        If InStr(htmlLiks.innerText, "糸満市") > 0 Then
'            糸満市だったらファイルダウンロードしてみる
'            リンクをクリック
'            htmlLiks.Click
'            ieObject.Visible = True
'            Call ShowWindow(ieObject.hwnd, SW_MINIMIZE)
'            Set fsoLink = New FileSystemObject
'            Dim strFilePath As String
'            ファイル名生成､Tempディレクトリで､拡張子は無しで設定する
'            strFilePath = fsoLink.BuildPath(fsoLink.GetSpecialFolder(TemporaryFolder), Format(Now(), "yyyymmddhhmmss"))
'            名前を付けて保存を実行､保存後のフルパス名が戻り値として返ってくる (多分拡張子とか付けてくれてるはず)
'            Dim strResultFilePath As String
'            strResultFilePath = IE_Save_As.DownloadNotificationBarSaveAs(ieObject.hwnd, strFilePath)
'            Call ieObject.Quit
'            Set ieObject = Nothing
'            If fsoLink.FileExists(strResultFilePath) Then
'                Application.StatusBar = strResultFilePath & " のダウンロード完了"
'            End If
'            Exit For
'        End If
'    Next htmlLiks
''-----------------------------------------------------------------------------------------------------------------------------
'
'    'とりあえずシートにキャラ名とかを出してみる
'    Application.ScreenUpdating = False
'    longColumnCount = START_COLUMN
'    longRowCount = START_ROW
'    'ヘッダ
'    For Each tableHeader In returnHTML.getElementsByName("sortabletable1")(0).getElementsByTagName("thead")(0).getElementsByTagName("th")
'        shIETest.Cells(longRowCount, longColumnCount).Value = tableHeader.innerText
'        '次の列へ
'        Debug.Print tableHeader.innerText
'        longColumnCount = longColumnCount + 1
'    Next tableHeader
'    'データ
'    longRowCount = START_ROW + 1
'    Application.StatusBar = "情報取得中..."
'    For Each tableRow In returnHTML.getElementsByName("sortabletable1")(0).getElementsByTagName("tbody")(0).getElementsByTagName("tr")
'        '各行に対して処理を行っていく
'        'trタグの中のtdタグの中身だけ引っこ抜く
'        '列を初期位置へ
'        longColumnCount = START_COLUMN
'        For Each tableData In tableRow.getElementsByTagName("td")
'            shIETest.Cells(longRowCount, longColumnCount).Value = tableData.innerText
''            Debug.Print tableData.innerHTML
'            '次の列へ
'            longColumnCount = longColumnCount + 1
'        Next tableData
'        '次の行へ
'        longRowCount = longRowCount + 1
'        Application.StatusBar = "情報取得中..." & longRowCount - START_ROW & " 件取得済み"
'        DoEvents
'    Next tableRow
'    Application.StatusBar = "取得完了"
    Application.ScreenUpdating = True
    Stop
    If Not ieObject Is Nothing Then
        Call ieObject.Quit
        Set ieObject = Nothing
    End If
End Function
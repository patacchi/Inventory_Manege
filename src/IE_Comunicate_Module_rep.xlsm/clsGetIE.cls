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
'----------------------------------------------------------------------------------------------------------------
'SaveAsの定数定義スタート
'Sleep Default Time
Private Const SLEEP_DEFAULT_MILLISEC As Long = 500                          'ループ待ちの標準ミリ秒
Private Const WM_SYSCHAR As Long = &H106&                                   'PostMessage用、System Charactor ALTキー押した状態のキーメッセージ
Private Const NOTIFICATION_CLASS_NAME As String = "Frame Notification Bar"  'FindWindow用、通知バーのクラス名
Private Const NOTIFICATION_SAVE_BUTTON_NAME As String = "保存"              '通知バーの[保存]ボタンのName
Private Const NOTIFICATION_TEXT As String = "通知バーのテキスト"            '通知バーのテキストのName
Private Const NOTIFICATION_CLOSE_BUTTON_NAME As String = "閉じる"           '通知バーの[閉じる]ボタンのName
Private Const ROLE_SYSTEM_BUTTONDROPDOWN = &H38&                            'ドロップダウンボタン
Private Const CONTEXT_MENU_CLASS_NAME As String = "#32768"                  'コンテキストメニューのクラス名、名前を付けて保存のドロップダウンした先のがこれ
Private Const SAVEASDIALOG_NAME As String = "名前を付けて保存"              '名前を付けて保存ダイアログボックスのName
Private Const SAVEASDIALOG_FILE_NAME As String = "ファイル名:"              'ファイル名コンボボックスのName
Private Const SAVEASDIALOG_SAVEAS_BUTTON_NAME As String = "保存(S)"         '保存ボタンのName
Private Const SAVEASTIMEOUT As Long = 30                                    'SaveAsウィンドウを検索するときのタイムアウト時間
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
'---------------------------------------------------------------------------------------------------------------
'SaveAs Modules
'''IEのハンドルと保存ファイル名を引数として、名前を付けて保存のボタン操作をする
'Return 最終的に保存されたフルパスをStringで返す
'注！指定されたファイルが存在する場合は強制的に削除されます！！！
Public Function DownloadNotificationBarSaveAs(ByVal strargSaveFilePath As String, Optional ByRef hIE As LongPtr) As String
    If hIE = 0 Then
        If Not ieObject Is Nothing Then
            hIE = ieObject.hwnd
        End If
    End If
    'ハンドルと保存ファイル名は必須なので、どちらかが指定されていなかったら抜ける
    If hIE = 0 Or strargSaveFilePath = "" Then
        Debug.Print Now() & " DownloadNotificationBarSaveAs: handle or Save Filneme is empty"
        DownloadNotificationBarSaveAs = ""
        Exit Function
    End If
    On Error GoTo ErrorCatch
    '指定されたファイルが存在していたら削除する
    With CreateObject("Scripting.FileSystemObject")
      If .FileExists(strargSaveFilePath) Then .DeleteFile strargSaveFilePath, True
    End With
    Dim uiAuto As CUIAutomation
    Set uiAuto = New CUIAutomation
    '通知バーの[別名で保存]を押す
    If Not PressSaveAsNotificationBar(uiAuto, hIE) Then
        '失敗したようですね
        'エラー表示とかは各プロシージャに記述する
        DownloadNotificationBarSaveAs = ""
        Exit Function
    End If
    '[名前を付けて保存]ダイアログ操作
    If Not SaveAsFilenameDialog(uiAuto, strargSaveFilePath) Then
        DownloadNotificationBarSaveAs = ""
        Exit Function
    End If
    'IEを通常表示に戻す
    Call ShowWindow(hIE, SW_RESTORE)
    'ダウンロード完了後通知バーを閉じる
    '戻り値としてファイル名が来てるはずなので、引数よりディレクトリ名を付加し、フルパスを返す
    Dim strResultFileName As String
    strResultFileName = ClosingNotificationBar(uiAuto, hIE)
    With CreateObject("Scripting.FilesystemObject")
        'ディレクトリ名を付加し、フルパスにする
        DownloadNotificationBarSaveAs = .BuildPath(.GetParentFolderName(strargSaveFilePath), strResultFileName)
        'デバッグ出力
        Debug.Print "DownloadNotificationBarSaveAs complete. file name: " & .BuildPath(.GetParentFolderName(strargSaveFilePath), strResultFileName)
    End With
    Set uiAuto = Nothing
    Exit Function
ErrorCatch:
    Debug.Print Now() & " DownloadNotificationBarSaveAs code: " & Err.Number & " Description: " & Err.Description
    DownloadNotificationBarSaveAs = ""
    Exit Function
End Function
'---------------------------------------------------------------------------------------------------------------
'''通知バーの[別名で保存]を押す
Private Function PressSaveAsNotificationBar(ByRef argUIAuto As CUIAutomation, ByVal hIEWnd As LongPtr) As Boolean
    'uiAutoとIEハンドル両方とも必須なので引数チェックし無かったら抜ける
    If argUIAuto Is Nothing Or hIEWnd = 0 Then
        Debug.Print Now() & " PressSaverAsNothificationBar: uiAuto or IEHandle empry"
        PressSaveAsNotificationBar = False
        Exit Function
    End If
    '通知バーを取得
    Dim hWndNotification As LongPtr
    Dim dateStart As Date
    '処理開始時間を取得
    dateStart = Now()
    Application.StatusBar = "通知バー取得中"
    On Error GoTo ErrorCatch
    Do
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        hWndNotification = FindWindowEx(hIEWnd, 0&, NOTIFICATION_CLASS_NAME, vbNullString)
        'タイムアウト時間を過ぎていたら強制終了する
        If Second(Now() - dateStart) >= SAVEASTIMEOUT Then
            MsgBox "通知バー検索時に" & SAVEASTIMEOUT & "秒のタイムアウト時間を超過しました。処理を中断します"
            PressSaveAsNotificationBar = False
            Exit Function
        End If
    Loop Until hWndNotification
    '通知バーが可視状態になるまで待機（これをやらないと操作に失敗することがある・・・らしい）
    Application.StatusBar = "通知バー準備中"
    Do
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        'タイムアウト時間を過ぎていたら強制終了する
        If Second(Now() - dateStart) >= SAVEASTIMEOUT Then
            MsgBox "通知バー検索時に" & SAVEASTIMEOUT & "秒のタイムアウト時間を超過しました。処理を中断します"
            PressSaveAsNotificationBar = False
            Exit Function
        End If
    Loop Until IsWindowVisible(hWndNotification)
    Debug.Print Second(Now() - dateStart) & " 秒で通知バー取得完了"
    '[保存]スプリットボタン取得
'    Application.StatusBar = "保存ボタン取得中"
    Dim elmNotificationBar As IUIAutomationElement
    Set elmNotificationBar = argUIAuto.ElementFromHandle(ByVal hWndNotification)
    Dim elmASaveSplitButton As IUIAutomationElement
    Do
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        Set elmASaveSplitButton = GetUIElement(argUIAuto, _
                                                elmNotificationBar, UIA_NamePropertyId, _
                                                NOTIFICATION_SAVE_BUTTON_NAME, _
                                                UIA_SplitButtonControlTypeId)
    Loop While elmASaveSplitButton Is Nothing
    '[保存]ボタンのドロップダウン取得
    Dim elmSaveAsDropDownButton As IUIAutomationElement
    Do
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        Set elmSaveAsDropDownButton = GetUIElement(argUIAuto, _
                                                    elmNotificationBar, _
                                                    UIA_LegacyIAccessibleRolePropertyId, _
                                                    ROLE_SYSTEM_BUTTONDROPDOWN, _
                                                    UIA_SplitButtonControlTypeId)
    Loop While elmSaveAsDropDownButton Is Nothing
    '[保存]ドロップダウンボタン押下
    Dim iptn As IUIAutomationInvokePattern
    Set iptn = elmSaveAsDropDownButton.GetCurrentPattern(UIA_InvokePatternId)
    'メニューウィンドウ（コンテキストメニュー）の取得
    Application.StatusBar = "コンテキストメニュー取得"
    Dim elmSaveMenyu As IUIAutomationElement
    Do
        'ドロップダウンボタン.click()
        iptn.Invoke
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        Set elmSaveMenyu = GetUIElement(argUIAuto, _
                                        argUIAuto.GetRootElement, _
                                        UIA_ClassNamePropertyId, _
                                        CONTEXT_MENU_CLASS_NAME, _
                                        UIA_MenuControlTypeId)
    Loop While elmSaveMenyu Is Nothing
    '[名前を付けて保存(A)]ボタン押下
    'ここだけPostMessageで行う
    Dim hWndSaveMenu As LongPtr
    hWndSaveMenu = FindWindow(CONTEXT_MENU_CLASS_NAME, vbNullString)
    PostMessage hWndSaveMenu, WM_SYSCHAR, vbKeyA, 0&
    PressSaveAsNotificationBar = True
    Exit Function
ErrorCatch:
    Debug.Print "PressSaveAsNothificationBar code: " & Err.Number & " Description: = " & Err.Description
    PressSaveAsNotificationBar = False
    Exit Function
End Function
'---------------------------------------------------------------------------------------------------------------
'''[名前を付けて保存]ダイアログボックスの操作
Private Function SaveAsFilenameDialog(ByRef argUIAuto As CUIAutomation, ByVal strSaveFilePath As String) As Boolean
    If argUIAuto Is Nothing Or strSaveFilePath = "" Then
        '引数が空だったら抜ける
        Debug.Print "SaveAsFilenameDialog: UIAuto or SaveFilePath is empty"
        SaveAsFilenameDialog = False
        Exit Function
    End If
    '[名前を付けて保存]ダイアログボックスの取得
    Application.StatusBar = "名前を付けて保存ダイアログボックス操作中"
    On Error GoTo ErrorCatch
    Dim elmSaveAsDialog As IUIAutomationElement
    Do
        DoEvents
        Set elmSaveAsDialog = GetUIElement(argUIAuto, _
                                            argUIAuto.GetRootElement, _
                                            UIA_NamePropertyId, _
                                            SAVEASDIALOG_NAME, _
                                            UIA_WindowControlTypeId)
        Sleep (SLEEP_DEFAULT_MILLISEC)
    Loop While elmSaveAsDialog Is Nothing
    'ファイル名エディットボックス（コンボボックス）の取得
    Dim elmFileNameComboBox As IUIAutomationElement
    Do
        Set elmFileNameComboBox = GetUIElement(argUIAuto, _
                                                elmSaveAsDialog, _
                                                UIA_NamePropertyId, _
                                                SAVEASDIALOG_FILE_NAME, _
                                                UIA_EditControlTypeId)
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
    Loop While elmFileNameComboBox Is Nothing
    'ファイルパスの入力
    Dim vptn As IUIAutomationValuePattern
    Set vptn = elmFileNameComboBox.GetCurrentPattern(UIA_ValuePatternId)
    vptn.SetValue strSaveFilePath
    '名前を付けて保存ダイアログの最小化を試みる
    Dim hWndSaveAsDialog As LongPtr
    hWndSaveAsDialog = FindWindow("#32770", SAVEASDIALOG_NAME)
    If hWndSaveAsDialog <> 0 Then
        '最小化してみる
        Call ShowWindow(hWndSaveAsDialog, SW_MINIMIZE)
    End If
    Application.StatusBar = "保存処理中"
    '[保存(S)]ボタン取得
    Dim elmSaveButton As IUIAutomationElement
    Do
        Set elmSaveButton = GetUIElement(argUIAuto, _
                                        elmSaveAsDialog, _
                                        UIA_NamePropertyId, _
                                        SAVEASDIALOG_SAVEAS_BUTTON_NAME, _
                                        UIA_ButtonControlTypeId)
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
    Loop While elmSaveButton Is Nothing
    '[保存]ボタン押下
    Dim iptn As IUIAutomationInvokePattern
    Set iptn = elmSaveButton.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke
    SaveAsFilenameDialog = True
    Exit Function
ErrorCatch:
    Debug.Print "SaveAsFilenameDialog code: " & Err.Number & " Description: " & Err.Description
    SaveAsFilenameDialog = False
    Exit Function
End Function
'---------------------------------------------------------------------------------------------------------------
'''ダウンロード完了後、通知バーを閉じる
Private Function ClosingNotificationBar(ByRef argUIAuto As CUIAutomation, ByVal hIEWnd As LongPtr) As String
    On Error GoTo ErrorCatch
    Application.StatusBar = "ダウンロード完了待ち"
    '通知バーを取得する
    Dim hWndNotification As LongPtr
    '処理開始時間を取得
    Dim dateStart As Date
    dateStart = Now()
    Do
        hWndNotification = FindWindowEx(hIEWnd, 0, NOTIFICATION_CLASS_NAME, vbNullString)
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        'タイムアウト時間を過ぎていたら強制終了する
        If Second(Now() - dateStart) >= SAVEASTIMEOUT Then
            MsgBox "通知バー検索時に" & SAVEASTIMEOUT & "秒のタイムアウト時間を超過しました。処理を中断します"
            Debug.Print "ClosingNotificationBar: NotificationBar find timeout"
            ClosingNotificationBar = False
            Exit Function
        End If
    Loop Until hWndNotification
    '可視状態になるまで待機
    Do
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
    Loop Until IsWindowVisible(hWndNotification)
    Dim elmNotificationBar As IUIAutomationElement
    Do
        Set elmNotificationBar = argUIAuto.ElementFromHandle(ByVal hWndNotification)
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
    Loop While elmNotificationBar Is Nothing
    '[通知バーのテキスト]取得
    Dim elmNotificationText As IUIAutomationElement
    Do
        Set elmNotificationText = GetUIElement(argUIAuto, _
                                                elmNotificationBar, _
                                                UIA_NamePropertyId, _
                                                NOTIFICATION_TEXT, _
                                                UIA_TextControlTypeId)
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        Application.StatusBar = "ダウンロード完了待ち " & Second(Now() - dateStart) & " 秒経過..."
    Loop While elmNotificationText Is Nothing
    '通知バーのテキストの内容を取得してみる
    Dim vptnNotificationText As IUIAutomationValuePattern
    Set vptnNotificationText = elmNotificationText.GetCurrentPattern(UIA_ValuePatternId)
    Do
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
        Application.StatusBar = vptnNotificationText.CurrentValue
    Loop While InStr(vptnNotificationText.CurrentValue, "ダウンロード済み") >= 1
    'テキスト出力してみる
    Dim strResultText As String
    strResultText = vptnNotificationText.CurrentValue
'    'デバッグ出力する
'    Debug.Print strResultText
    Dim arrResult() As String
    arrResult = Strings.Split(strResultText, " ")
    '[閉じる]ボタン取得
    Dim elmCloseButton As IUIAutomationElement
    Do
        Set elmCloseButton = GetUIElement(argUIAuto, _
                                            elmNotificationBar, _
                                            UIA_NamePropertyId, _
                                            NOTIFICATION_CLOSE_BUTTON_NAME, _
                                            UIA_ButtonControlTypeId)
        DoEvents
        Sleep (SLEEP_DEFAULT_MILLISEC)
    Loop While elmCloseButton Is Nothing
    '[閉じる]ボタン押下
    Dim iptnClose As IUIAutomationInvokePattern
    Set iptnClose = elmCloseButton.GetCurrentPattern(UIA_InvokePatternId)
    iptnClose.Invoke
    If UBound(arrResult) >= 1 Then
        ClosingNotificationBar = arrResult(0)
        Exit Function
    Else
        Debug.Print "ClosingNotificationBar: 結果のテキスト取得失敗してるっぽい"
        ClosingNotificationBar = ""
        Exit Function
    End If
ErrorCatch:
    Debug.Print "ClosingNotificationBar code: " & Err.Number & " Description: " & Err.Description
    ClosingNotificationBar = ""
    Exit Function
End Function
'---------------------------------------------------------------------------------------------------------------
'''uiAuto 指定されたプロパティID、コントロールタイプIDで指定された値を持つ要素を返す
'''return IUIAutomationElement
Private Function GetUIElement(ByVal uiAuto As CUIAutomation, _
                                ByVal elmParent As IUIAutomationElement, _
                                ByVal propertyID As Long, _
                                ByVal propertyValue As Variant, _
                                Optional ByVal ctrlType As Long = 0) _
                                As IUIAutomationElement
    '検索条件の設定
    Dim condFirst As IUIAutomationCondition
    Set condFirst = uiAuto.CreatePropertyCondition(propertyID, propertyValue)
    If ctrlType <> 0 Then
        'コントロールIDが指定されている場合は追加で以下を実行する
        Dim condSecond As IUIAutomationCondition
        Set condSecond = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, ctrlType)
        'propetryValue および ctrlID両方に一致する条件を作成する
        Set condFirst = uiAuto.CreateAndCondition(condFirst, condSecond)
    End If
    Set GetUIElement = elmParent.FindFirst(TreeScope_Subtree, condFirst)
End Function
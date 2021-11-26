Attribute VB_Name = "mobZaikoSerch"
Option Explicit
Private Const zaikoSerchURL As String = "http://www.freeway.fuchu.toshiba.co.jp/faz/zaikoSearch/"
Private Const DEBUG_SHOW_IE As Long = &H1                   'IEの画面を表示させるフラグ(1bit)
Private Const ZAIKO_SERCH_DL_TREE As String = "d1d0"        '在庫検索のダウンロードボタン（検索後のページ）がある階層文字列
'private const ZAIKO_SERCH_SCRIPT_TREE as String = "d1d0d"
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''IEから情報をとってきて（シートに書き出す）テストモジュール
'''--------------------------------------------------------------------------------------------------------------
Public Sub IETest()
    Dim getIETest As clsGetIE
    Set getIETest = New clsGetIE
    '在庫情報検索ページを設定
    getIETest.URL = zaikoSerchURL
'    getIETest.URL = "file:///C:/Users/q3005sbe/AppData/Local/Rep/Backup/FrameSampe.htm"
    'Debug用
'    isCollect = getIETest.OpenIEwithURL
    '指定したURLのHTML DocをDictionaryで受け取るテスト
    Dim longDebugFlag As Long                   'デバッグフラグを管理するためのLong変数
'    longDebugFlag = 0 Or DEBUG_SHOW_IE
    If longDebugFlag And DEBUG_SHOW_IE Then
        'IE表示フラグが立ってたのでプロパティ設定
        getIETest.Visible = True
    End If
    On Error Resume Next
    Dim dicReturnHTMLDoc As Dictionary
    Set dicReturnHTMLDoc = getIETest.ResultHTMLDoc
    If Err.Number <> 0 Then
        'エラー発生してたらとりあえずここに来てみる
        DebugMsgWithTime "IETest code: " & Err.Number & " Description: " & Err.Description
    End If
'    SetZaikoSerch_TehaiCode getIETest, InputBox("手配コードを入力して下さい")
    '検索する手配コードを一時的にshIEData.cells(4,3)に記入することとする
    SetZaikoSerch_TehaiCode getIETest, CStr(shIEData.Cells(4, 3).Value)
    Application.Wait 2
    '全フレームより指定したタグのHTML Documentを取ってくる
    Dim dicTagElms As Dictionary
    Set dicTagElms = getIETest.GetHTMLdicBydicHTMLDocandTagName(dicReturnHTMLDoc, "Input")
    'ダウンロードボタンを押してみる
    Dim docZaikoDLButton As HTMLDocument
    If dicReturnHTMLDoc.Exists(ZAIKO_SERCH_DL_TREE) Then
        '結果dicにダウンロードボタンの階層文字列がある場合のみ実行
        Set docZaikoDLButton = dicReturnHTMLDoc(ZAIKO_SERCH_DL_TREE)
    End If
    'ダウンロードボタンくりこ
    Dim docConfirm As HTMLDocument
    Set docConfirm = dicReturnHTMLDoc("d1d0d")
    docConfirm.parentWindow.execScript "chkSetChild( document );"
    docConfirm.parentWindow.execScript "$('#mainFm').attr('action', '../zaikoInfoSearch/validate/');"
    docConfirm.parentWindow.execScript "if(validateSearchCondition()) { document.forms[0].action = '../zaikoInfoSearch/download/'; document.forms[0].submit();}"
'    Stop
'    Sleep 3000
    '保存ファイル名生成
    Dim strFilePath As String
    Dim fsoLink As Scripting.FileSystemObject
    Set fsoLink = New FileSystemObject
    strFilePath = fsoLink.BuildPath(fsoLink.GetSpecialFolder(TemporaryFolder), Format(Now(), "yyyymmddhhmmss"))
    'SaveAs 操作
    Dim strResultFullPath As String
    strResultFullPath = getIETest.DownloadNotificationBarSaveAs(strFilePath, getIETest.IEInstance.hwnd)
    '帰ってきたBookを開いてみる
    getIETest.IEInstance.Visible = False
    Dim wkbNewBook As Workbook
    Set wkbNewBook = Workbooks.Open(strResultFullPath)
    wkbNewBook.Activate
'    '試しに検索ボタンをクリックしてみる
'    getIETest.IEInstance.document.frames(1).document.frames(0).document.getElementById("kensakuButton").Click
'    Dim localHTMLDoc As HTMLDocument
''    Set localHTMLDoc = dicReturnHTMLDoc(1).frames(0).document
'    Set localHTMLDoc = dicReturnHTMLDoc("t10")
'    Dim elementStrArray() As String
'    elementStrArray = getIETest.getTextArrayByTagName(localHTMLDoc, "A")
    Cells(getIETest.shRow, getIETest.shColumn).Value = dicReturnHTMLDoc(1).Title
    'Stop
    Set getIETest = Nothing
End Sub
'''Author Daisuke_Oota
'''GetIEクラスを引数として、在庫検索の手配コードに指定の文字列をセットする
'''
Private Sub SetZaikoSerch_TehaiCode(ByRef clsargIE As clsGetIE, strargTeheaiCode As String)
    If strargTeheaiCode = "" Then
        Exit Sub
    End If
    'IEのインスタンスに対して在庫検索の手配コードを設定してやる
    clsargIE.IEInstance.document.frames(1).document.frames(0).document.forms(0).Item(ZAIKO_SERCH_TEHAI_CODE_INPUT_BOX_NAME).Value = strargTeheaiCode
    '試験で表示させてみる
    Dim longDebugFrag As Long
    longDebugFrag = longDebugFrag Or DEBUG_SHOW_IE
    If longDebugFrag And DEBUG_SHOW_IE Then
        clsargIE.IEInstance.Visible = True
    End If
End Sub
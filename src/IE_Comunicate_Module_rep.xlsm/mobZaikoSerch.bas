Attribute VB_Name = "mobZaikoSerch"
Option Explicit
Private Const zaikoSerchURL As String = "http://www.freeway.fuchu.toshiba.co.jp/faz/zaikoSearch/"
Private Const DEBUG_SHOW_IE As Long = &H1                           'IEの画面を表示させるフラグ(1bit)
Private Const ZAIKO_SERCH_DL_TREE As String = "d1d0"                '在庫検索のダウンロードボタン（検索後のページ）がある階層文字列
Private Const ZAIKO_SERCH_SCRIPT_TREE As String = "d1d0d"           '在庫検索のページのスクリプト発動する階層文字列
Private Const ZAIKO_SERCH_DL_SCRIPT As String = "chkSetChild( document );$('#mainFm').attr('action', '../zaikoInfoSearch/validate/');if(validateSearchCondition()) { document.forms[0].action = '../zaikoInfoSearch/download/'; document.forms[0].submit();}"   '在庫検索のダウンロードボタンのスクリプト
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''手配コードを引数として、在庫検索（のファイルDL）を行うプロシージャ
'''--------------------------------------------------------------------------------------------------------------
Public Sub ZaikoSerchbyTehaiCode(ByVal strTehaiCode As String)
    Dim getIETest As clsGetIE
    Set getIETest = New clsGetIE
    If strTehaiCode = "" Then
        '手配コードが指定されていなかったら抜ける
        MsgBox "ZaikoSerchbyTehaiCode: 手配コードが空でした（必須項目）"
        Exit Sub
    End If
    '在庫情報検索ページを設定
    getIETest.URL = zaikoSerchURL
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
    '指定したURLより全フレームのHTMLDocを取得する Dictionary形式
    Set dicReturnHTMLDoc = getIETest.ResultHTMLDoc
    If Err.Number <> 0 Then
        'エラー発生してたらとりあえずここに来てみる
        DebugMsgWithTime "IETest code: " & Err.Number & " Description: " & Err.Description
    End If
    On Error GoTo ErrorCatch
    '検索する手配コードをセットしてやる
    SetZaikoSerch_TehaiCode getIETest, strTehaiCode
    'ダウンロードボタンくりこ
    'スクリプト直接実行に切り替え(confirm潰せなかった・・・）
    If dicReturnHTMLDoc.Exists(ZAIKO_SERCH_SCRIPT_TREE) Then
        '在庫検索スクリプトページの階層文字列が存在する場合のみ実行する
        Dim docConfirm As HTMLDocument
        Set docConfirm = dicReturnHTMLDoc(ZAIKO_SERCH_SCRIPT_TREE)
        docConfirm.parentWindow.execScript ZAIKO_SERCH_DL_SCRIPT
'        docConfirm.parentWindow.execScript "chkSetChild( document );"
'        docConfirm.parentWindow.execScript "$('#mainFm').attr('action', '../zaikoInfoSearch/validate/');"
'        docConfirm.parentWindow.execScript "if(validateSearchCondition()) { document.forms[0].action = '../zaikoInfoSearch/download/'; document.forms[0].submit();}"
    End If
    '-----------------------------------------------------------------------------------------------------------
    'Saveの場合（基本はこっち）
    '保存ファイル名の生成（ファイル名のみ、ディレクトリはDownloadの場所になるはずなので可変）
    Dim strFleName As String
    strFleName = "ZaikoSerch" & (Format(Now(), "yyyymmddhhmmss"))
    Dim strResultFilePath As String
    '保存ボタンを押し、結果のファイル名を受け取る
    strResultFilePath = getIETest.DownloadSave_NotificationBar(strFleName)
    'これで保存したファイル名がフルパスで取得できているので、あとは利用するのみ
'    MsgBox strResultFilePath
'    Call Application.Workbooks.Open(strResultFilePath)
'    '-----------------------------------------------------------------------------------------------------------
'    'SaveAsの時の使用方法
'    '保存ファイル名生成
'    Dim strFilePath As String
'    Dim fsoLink As Scripting.FileSystemObject
'    Set fsoLink = New FileSystemObject
'    strFilePath = fsoLink.BuildPath(fsoLink.GetSpecialFolder(TemporaryFolder), Format(Now(), "yyyymmddhhmmss"))
'    'SaveAs 操作
'    Dim strResultFullPath As String
'    strResultFullPath = getIETest.DownloadNotificationBarSaveAs(strFilePath, getIETest.IEInstance.Hwnd)
'    '帰ってきたBookを開いてみる
'    getIETest.IEInstance.Visible = False
'    Dim wkbNewBook As Workbook
'    Set wkbNewBook = Workbooks.Open(strResultFullPath)
'    wkbNewBook.Activate
'    '試しに検索ボタンをクリックしてみる
'    getIETest.IEInstance.document.frames(1).document.frames(0).document.getElementById("kensakuButton").Click
'    Dim localHTMLDoc As HTMLDocument
''    Set localHTMLDoc = dicReturnHTMLDoc(1).frames(0).document
'    Set localHTMLDoc = dicReturnHTMLDoc("t10")
'    Dim elementStrArray() As String
'    elementStrArray = getIETest.getTextArrayByTagName(localHTMLDoc, "A")
'    Cells(getIETest.shRow, getIETest.shColumn).Value = dicReturnHTMLDoc(1).Title
    'Stop
    Set getIETest = Nothing
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "ZaikoSerchbyTehaiCode code: " & Err.Number & " Description: " & Err.Description
    If Not getIETest Is Nothing Then
        Set getIETest = Nothing
    End If
    Exit Sub
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
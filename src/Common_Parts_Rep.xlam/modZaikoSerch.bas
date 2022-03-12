Attribute VB_Name = "modZaikoSerch"
Option Explicit
Private Const zaikoSerchURL As String = "http://www.freeway.fuchu.toshiba.co.jp/faz/zaikoSearch/"
Private Const DEBUG_SHOW_IE As Long = &H1                           'IEの画面を表示させるフラグ(1bit)
Private Const ZAIKO_SERCH_DL_TREE As String = "d1d0"                '在庫検索のダウンロードボタン（検索後のページ）がある階層文字列
Private Const ZAIKO_SERCH_SCRIPT_TREE As String = "d1d0d"           '在庫検索のページのスクリプト発動する階層文字列
Private Const ZAIKO_SERCH_DL_SCRIPT As String = "chkSetChild( document );$('#mainFm').attr('action', '../zaikoInfoSearch/validate/');if(validateSearchCondition()) { document.forms[0].action = '../zaikoInfoSearch/download/'; document.forms[0].submit();}"   '在庫検索のダウンロードボタンのスクリプト
'''Author Daisuke_Oota
'''--------------------------------------------------------------------------------------------------------------
'''Summary
'''手配コードを引数として、在庫検索（のファイルDL）を行い、DLしたファイル名のフルパスを返す
'''戻り値 string    DLしたファイルのフルパス
'''parms
'''IEZaikoSerch                     インスタンスを共有してコンストラクタでもたつくのを解消したい
'''optional strargReturnFileName
'''--------------------------------------------------------------------------------------------------------------
Public Function ZaikoSerchbyTehaiCode(ByVal strTehaiCode As String, _
    ByRef clsGetieZaikoSerch As clsGetIE, Optional strargReturnFileName As String) As String
    'クラス引数確認
    If clsGetieZaikoSerch Is Nothing Then
        'クラス引数が初期化されていない
        DebugMsgWithTime "ZaikoSerhbyTehaiCode: Warning! clsGetIE instance empy. will delay...."
        Set clsGetieZaikoSerch = New clsGetIE
    End If
    '手配コード空文字はありえるので続行する
'    If strTehaiCode = "" Then
'        '手配コードが指定されていなかったら抜ける
'        MsgBox "ZaikoSerchbyTehaiCode: 手配コードが空でした（必須項目）"
'        Exit Function
'    End If
    '在庫情報検索ページを設定
    clsGetieZaikoSerch.URL = zaikoSerchURL
    Dim longDebugFlag As Long                   'デバッグフラグを管理するためのLong変数
'    longDebugFlag = 0 Or DEBUG_SHOW_IE
'    If longDebugFlag And DEBUG_SHOW_IE Then
'        'IE表示フラグが立ってたのでプロパティ設定
'        clsGetieZaikoSerch.Visible = True
'    End If
    On Error Resume Next
    Dim dicReturnHTMLDoc As Dictionary
    If Not dicReturnHTMLDoc Is Nothing Then
        '2週目以降はインスタンス再利用するため、Dictionaryに中身が入ったままになっている
        'RemoveAllを試してみる
        'ダメだった場合は１週ごとにNothingにするように
        dicReturnHTMLDoc.RemoveAll
    End If
    '指定したURLより全フレームのHTMLDocを取得する Dictionary形式
    Set dicReturnHTMLDoc = clsGetieZaikoSerch.ResultHTMLDoc
    If Err.Number <> 0 Then
        'エラー発生してたらとりあえずここに来てみる
        DebugMsgWithTime "ZaikoSerchbyTehaiCode code: " & Err.Number & " Description: " & Err.Description
    End If
    On Error GoTo ErrorCatch
    '検索する手配コードをセットしてやる
    SetZaikoSerch_TehaiCode clsGetieZaikoSerch, strTehaiCode
    'ダウンロードボタンくりこ
    'スクリプト直接実行に切り替え(confirm潰せなかった・・・）
    If dicReturnHTMLDoc.Exists(ZAIKO_SERCH_SCRIPT_TREE) Then
        '在庫検索スクリプトページの階層文字列が存在する場合のみ実行する
        Dim docConfirm As HTMLDocument
        If Not docConfirm Is Nothing Then
            'この時点でdocConfirmがNothingじゃなかった場合
'            docConfirm.Close
            docConfirm.Clear
        End If
        Set docConfirm = dicReturnHTMLDoc(ZAIKO_SERCH_SCRIPT_TREE)
        docConfirm.parentWindow.execScript ZAIKO_SERCH_DL_SCRIPT
    End If
    '-----------------------------------------------------------------------------------------------------------
    'Saveの場合（基本はこっち）
    '保存ファイル名の生成（ファイル名のみ、ディレクトリはDownloadの場所になるはずなので可変）
    If strargReturnFileName = "" Then
        '保存ファイル名が指定されなかった場合
        'TehaiCode_yyyy_mm_dd_HH_MM_SS_fff
        strargReturnFileName = strTehaiCode & GetTimeForFileNameWithMilliSec
    End If
    Dim strResultFilePath As String
    '保存ボタンを押し、結果のファイル名を受け取る
    strResultFilePath = clsGetieZaikoSerch.DownloadSave_NotificationBar(strargReturnFileName)
    ZaikoSerchbyTehaiCode = strResultFilePath
'    '-----------------------------------------------------------------------------------------------------------
'    'SaveAsの時の使用方法
'    '保存ファイル名生成
'    Dim strFilePath As String
'    Dim fsoLink As Scripting.FileSystemObject
'    Set fsoLink = New FileSystemObject
'    strFilePath = fsoLink.BuildPath(fsoLink.GetSpecialFolder(TemporaryFolder), Format(Now(), "yyyymmddhhmmss"))
'    'SaveAs 操作
'    Dim strResultFullPath As String
'    strResultFullPath = clsgetiezaikoserch.DownloadNotificationBarSaveAs(strFilePath, clsgetiezaikoserch.IEInstance.Hwnd)
'    '帰ってきたBookを開いてみる
'    clsgetiezaikoserch.IEInstance.Visible = False
'    Dim wkbNewBook As Workbook
'    Set wkbNewBook = Workbooks.Open(strResultFullPath)
'    wkbNewBook.Activate
'    '試しに検索ボタンをクリックしてみる
'    clsgetiezaikoserch.IEInstance.Document.frames(1).Document.frames(0).Document.getElementById("kensakuButton").Click
'    Dim localHTMLDoc As HTMLDocument
''    Set localHTMLDoc = dicReturnHTMLDoc(1).frames(0).Document
'    Set localHTMLDoc = dicReturnHTMLDoc("t10")
'    Dim elementStrArray() As String
'    elementStrArray = clsgetiezaikoserch.getTextArrayByTagName(localHTMLDoc, "A")
'    Cells(clsgetiezaikoserch.shRow, clsgetiezaikoserch.shColumn).Value = dicReturnHTMLDoc(1).Title
'    Set clsGetieZaikoSerch = Nothing
    Exit Function
ErrorCatch:
    DebugMsgWithTime "ZaikoSerchbyTehaiCode code: " & Err.Number & " Description: " & Err.Description
    'クラス変数はインスタンスを共有するので個別に解放はNG
'    If Not clsGetieZaikoSerch Is Nothing Then
'        Set clsGetieZaikoSerch = Nothing
'    End If
    Exit Function
End Function
'''Author Daisuke_Oota
'''GetIEクラスを引数として、在庫検索の手配コードに指定の文字列をセットする
'''args
Private Sub SetZaikoSerch_TehaiCode(ByRef clsargIE As clsGetIE, strargTeheaiCode As String)
    On Error GoTo ErrorCatch
    If strargTeheaiCode = "" Then
'        Exit Sub
        '管理課指定した上での手配コード空白はあり得るので、ダイアログを出して処理を分岐する
        Dim resultFullDL As VbMsgBoxResult
        resultFullDL = MsgBox("手配コードが指定されませんでした。全ての手配コードのファイルをDLしますか？", vbYesNo)
        If resultFullDL = vbNo Then
            'NO、いいえが押された
            MsgBox "処理を中断します"
            Exit Sub
        End If
    End If
    'IEインスタンス（在庫検索ページ）の管理課に対して「W」を設定してやる
    '現状 Index = 11 が MSブ W なのでそこを選択してやる、画面上の表示は変わっていないが、データ上は反映されている模様
    clsargIE.IEInstance.Document.frames(1).Document.frames(0).Document.forms(0).Item(ZAIKO_SERTH_KANRI_KA_INPUT_BOX_NAME).selectedIndex = 11
    'IEのインスタンスに対して在庫検索の手配コードを設定してやる
    clsargIE.IEInstance.Document.frames(1).Document.frames(0).Document.forms(0).Item(ZAIKO_SERCH_TEHAI_CODE_INPUT_BOX_NAME).Value = strargTeheaiCode
#If DebugShowIE Then
    '条件付きコンパイル引数で表示する設定になっていたら表示してやる
    clsargIE.IEInstance.Visible = True
#End If
Exit Sub
ErrorCatch:
    DebugMsgWithTime "SetZaikoSerch_TehaiCode: " & Err.Number & " Description: " & Err.Description
    Exit Sub
End Sub
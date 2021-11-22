Attribute VB_Name = "PublicTestGetIE"
Option Explicit
Private Const zaikoSerchURL As String = "http://www.freeway.fuchu.toshiba.co.jp/faz/zaikoSearch/"
Private Const DEBUG_SHOW_IE As Long = &H1                   'IEの画面を表示させるフラグ(1bit)
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
    Set dicReturnHTMLDoc = getIETest.ReturnHTMLDocbyURL
    If Err.Number <> 0 Then
        'エラー発生してたらとりあえずここに来てみる
        Stop
    End If
    SetZaikoSerch_TehaiCode getIETest, InputBox("手配コードを入力して下さい")
    Application.Wait 2
    '試しに検索ボタンをクリックしてみる
    getIETest.IEInstance.document.frames(1).document.frames(0).document.getElementById("kensakuButton").Click
    Dim localHTMLDoc As HTMLDocument
    Set localHTMLDoc = dicReturnHTMLDoc(1).frames(0).document
    Dim elementStrArray() As String
    elementStrArray = getIETest.getTextArrayByTagName(localHTMLDoc, "A")
    Cells(getIETest.shRow, getIETest.shColumn).Value = dicReturnHTMLDoc(1).Title
    Stop
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
Attribute VB_Name = "Mod_WinAPI"
Option Explicit
'WindowAPI 関数宣言
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, _
    ByVal nWidth As Long, _
    ByVal nEscapement As Long, _
    ByVal nOrientation As Long, _
    ByVal fnWeight As Long, _
    ByVal IfdwItalic As Long, _
    ByVal fdwUnderline As Long, _
    ByVal fdwStrikeOut As Long, _
    ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, _
    ByVal fdwClipPrecision As Long, _
    ByVal fdwQuality As Long, _
    ByVal fdwPitchAndFamily As Long, _
    ByVal lpszFace As String) As LongPtr
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long
'---------------------------------------------------------------------
'定数・構造体宣言
'RECT構造体定義
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'CreateFont用定数
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_SCRIPT = 64
Private Const DT_CALCRECT = &H400
'-----------------------------------------------------------------------------------------------------------------------
'プロシージャ定義
'''与えらた文字列、フォント、サイズから画面の実描画長を取得する関数
'''Return Long      実描画ポイント数?
'''args
'''strargTargetText         ターゲットとなるテキスト
'''strargFONT_NAME          フォント名
'''longargFont_Height       フォントサイズ?縦の長さらしいけど・・・
'''Optional longargFontWidtsScal   フォントの横拡大サイズを%で デフォルトは100
Public Function MesureTextWidth( _
    strargTargetText As String, _
    strargFONT_NAME As String, _
    longargFont_Height As Long, _
    Optional longargFontWidtsScale = 100) As Long
    On Error GoTo ErrorCatch
    '画面全体の描画領域のハンドルを取得
    Dim hwholeScreenDC As LongPtr
    hwholeScreenDC = GetDC(0&)
    '仮想画面描画領域のハンドルを取得
    Dim hvirtualDC As LongPtr
    hvirtualDC = CreateCompatibleDC(hwholeScreenDC)
    '文字の拡大率に応じて横幅を設定
    Dim longWidth As Long
    Select Case longargFontWidtsScale
    Case 100
        '0、等倍の場合
        '自動設定にする・・・とうまくいかない・・？
        'とりあえず高さ(Font.Size)と同じにしてみる
        longWidth = longargFont_Height
    Case Else
        '倍率指定されていた場合
        '指定倍率を掛けてやる
        '実測値とちょっと違う結果になってたので後ろで補正
        longWidth = CLng(longargFont_Height * (longargFontWidtsScale / 100) * 80 / 100)
    End Select
    'フォントのハンドルを取得
    Dim hFont As LongPtr
    hFont = CreateFont(longargFont_Height, longWidth, 0, 0, FW_NORMAL, _
    0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
    CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, _
    DEFAULT_PITCH Or FF_SCRIPT, strargFONT_NAME)
    '仮想描画領域にフォントを設定
    Call SelectObject(hvirtualDC, hFont)
    '描画領域の周囲を取得
    Dim DrawAreaRectangle As RECT
    '設定したフォントを適用し、テキスト書き出し
    Call DrawText(hvirtualDC, strargTargetText, -1, DrawAreaRectangle, DT_CALCRECT)
    '使用したオブジェクトを開放する
    Call DeleteObject(hFont)
    Call DeleteObject(hvirtualDC)
    Call ReleaseDC(0&, hwholeScreenDC)
    '結果を返す
    MesureTextWidth = DrawAreaRectangle.Right - DrawAreaRectangle.Left
    GoTo CloseAndExit
ErrorCatch:
    Debug.Print "MesureTextWidth code : " & Err.Number & " Descriptoin: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
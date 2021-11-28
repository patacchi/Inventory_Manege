Attribute VB_Name = "modWinAPI"
Option Explicit
'Windows API 関数定義
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndparent As LongPtr, ByVal hWndchild As LongPtr, ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetClassLongPtr Lib "user32" Alias "SetClassLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndparent As Long, ByVal hWndchild As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwndOwner As LongPtr, ByVal wCmdGW_ As Long) As LongPtr
'-----------------------------------------------------------------------------------------------------------------------
'UNC対応のため、Win32API使用
Public Declare PtrSafe Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As LongPtr
'-----------------------------------------------------------------------------------------------------------------------
'定数・構造体定義
Public Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
'日付をミリ秒単位で取得するのにWin32APIを使用
'SYSTEMTIME構造体定義
Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
'const GetWindowLongPtr() 及び SetWindowLongPtr()で使用する定数
Public Const GWL_STYLE As Long = (-16)              'ウィンドウスタイルのハンドラ番号
Public Const WS_MAXIMIZEBOX As Long = &H10000       'ウィンドウスタイルで最大化ボタンをつける
Public Const WS_MINIMIZEBOX As Long = &H20000       'ウィンドウスタイルで最小化ボタンを付ける
Public Const WS_THICKFRAME As Long = &H40000        'ウィンドウスタイルでサイズ変更をつける
Public Const WS_SYSMENU As Long = &H80000           'ウィンドウスタイルでコントロールメニューボックスをもつウィンドウを作成する
' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10
'GetWindows() GW_Cmd
Public Const GW_HWNDFIRST As Long = 0
Public Const GW_HWNDLAST As Long = 1
Public Const GW_HWNDNEXT As Long = 2
Public Const GW_HWNDPREV As Long = 3
Public Const GW_OWNER As Long = 4
Public Const GW_CHILD As Long = 5
Public Const GW_ENABLEPOPUP As Long = 6
'-----------------------------------------------------------------------------------------------------------------------
'プロシージャ定義
'フォームに最大化・リサイズ機能を追加する。
Public Sub FormResize(Optional hwnd As LongPtr = 0)
    Dim WndStyle As LongPtr
    If hwnd = 0 Then
        'ハンドルが指定されなかった場合は、アクティブウィンドウのハンドルを取得する
        'ウィンドウハンドルの取得
        hwnd = GetActiveWindow()
    End If
    'ウィンドウのスタイルを取得
    WndStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
    '最大・最小・サイズ変更を追加する
    WndStyle = WndStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_SYSMENU
    Call SetWindowLongPtr(hwnd, GWL_STYLE, WndStyle)
End Sub
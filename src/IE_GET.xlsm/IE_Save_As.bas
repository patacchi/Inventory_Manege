Attribute VB_Name = "IE_Save_As"
'''IE操作において、ファイルダウンロード時に名前を付けて保存（Save as)を行うためのモジュール
'参照設定： UIAutomationClient
Option Explicit
'''Windows API宣言
#If VBA7 And Win64 Then
'    Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndparent As LongPtr, ByVal hWndchild As LongPtr, ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndparent As Long, ByVal hWndchild As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
'定数宣言
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
'Sleep Default Time
Public Const SLEEP_DEFAULT_MILLISEC As Long = 100                           'ループ待ちの標準ミリ秒
Public Const WM_SYSCHAR As Long = &H106&                                    'PostMessage用、System Charactor ALTキー押した状態のキーメッセージ
Public Const NOTIFICATION_CLASS_NAME As String = "Frame Notification Bar"   'FindWindow用、通知バーのクラス名
Public Const NOTIFICATION_SAVE_BUTTON_NAME As String = "保存"               '通知バーの[保存]ボタンのName
Public Const NOTIFICATION_TEXT As String = "通知バーのテキスト"             '通知バーのテキストのName
Public Const NOTIFICATION_CLOSE_BUTTON_NAME As String = "閉じる"            '通知バーの[閉じる]ボタンのName
Public Const ROLE_SYSTEM_BUTTONDROPDOWN = &H38&                             'ドロップダウンボタン
Public Const CONTEXT_MENU_CLASS_NAME As String = "#32768"                   'コンテキストメニューのクラス名、名前を付けて保存のドロップダウンした先のがこれ
Public Const SAVEASDIALOG_NAME As String = "名前を付けて保存"               '名前を付けて保存ダイアログボックスのName
Public Const SAVEASDIALOG_FILE_NAME As String = "ファイル名:"               'ファイル名コンボボックスのName
Public Const SAVEASDIALOG_SAVEAS_BUTTON_NAME As String = "保存(S)"          '保存ボタンのName
Public Const SAVEASTIMEOUT As Long = 30                                     'SaveAsウィンドウを検索するときのタイムアウト時間

'''Module
'---------------------------------------------------------------------------------------------------------------
'''IEのハンドルと保存ファイル名を引数として、名前を付けて保存のボタン操作をする
'Return 最終的に保存されたファイル名(パス無し）をStringで返す
'注！指定されたファイルが存在する場合は強制的に削除されます！！！
Public Function DownloadNotificationBarSaveAs(ByRef hIE As LongPtr, ByVal strargSaveFilePath As String) As String
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
        Set elmSaveAsDialog = GetUIElement(argUIAuto, _
                                            argUIAuto.GetRootElement, _
                                            UIA_NamePropertyId, _
                                            SAVEASDIALOG_NAME, _
                                            UIA_WindowControlTypeId)
        DoEvents
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

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputStocINDate 
   Caption         =   "入庫日入力フォーム"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   OleObjectBlob   =   "frmInputStocINDate.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmInputStocINDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''-------------------------------------------------------------------------------------------
'''イベント
'''UserFormInitialize
Private Sub UserForm_Initialize()
    ConstRuctor
End Sub
'''UserFormTerminate
'''ModalessにしてるのでTerminateで確実にUnloadする事
Private Sub UserForm_Terminate()
    DestRuctor
End Sub
'''入庫日テキストボックスMouseUpイベント、このテキストボックスは基本的にはDatePickerから選んでもらうのでユーザーは入力しない
Private Sub txtBoxStockINDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim dateStockIN As Date
    dateStockIN = PublicModule.GetDateUseDatePicker
    If CDbl(dateStockIN) = 0 Then
        '選択失敗してたら何もしない
        Exit Sub
    End If
    txtBoxStockINDate.Text = dateStockIN
    Exit Sub
End Sub
'''-------------------------------------------------------------------------------------------
'''メソッド
'''コンストラクタ
Private Sub ConstRuctor()
    'とりあえず入庫日を本日に
    txtBoxStockINDate.Text = Format(Year(Now()), "0000") & "/" & Format(Month(Now()), "00") & "/" & Format(Day(Now()), "00")
End Sub
'''デストラクタ
Private Sub DestRuctor()
    Unload Me
End Sub
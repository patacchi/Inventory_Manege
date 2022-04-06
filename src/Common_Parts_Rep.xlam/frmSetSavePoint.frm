VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetSavePoint 
   Caption         =   "印刷リスト識別名入力画面"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4395
   OleObjectBlob   =   "frmSetSavePoint.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSetSavePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''binLabelのSavePointを入力するフォーム
Option Explicit
Private Const SAVEPOINT_NAME_1 As String = "BINカードラベル出力用"
Private Const SAVEPOINT_NAME_2 As String = "入庫用"
Private Const SAVEPOINT_NAME_3 As String = "出庫用"
Private Const SAVEPOINT_NAME_4 As String = "現品票出力用"
Private Const SAVEPOINT_NAME_5 As String = "スペック表 小 (詳細現品票)"
Private Sub UserForm_Initialize()
    ConstRuctor
End Sub
'click
'入力完了
Private Sub btnCompInput_Click()
    'FormCommonのグローバル変数に結果を格納し、自身はUnload
    FormCommon.strSavePointName = cmbBox_SavePointName.Text
    Unload Me
End Sub
'''フォームコンストラクタ
Private Sub ConstRuctor()
    'グローバル変数の結果格納用変数を空文字にリセットする
    FormCommon.strSavePointName = ""
    'コンボボックスに定型文を設定
    cmbBox_SavePointName.AddItem SAVEPOINT_NAME_1
    cmbBox_SavePointName.AddItem SAVEPOINT_NAME_2
    cmbBox_SavePointName.AddItem SAVEPOINT_NAME_3
    cmbBox_SavePointName.AddItem SAVEPOINT_NAME_4
    cmbBox_SavePointName.AddItem SAVEPOINT_NAME_5
    'コンボボックス初期値は空文字
    cmbBox_SavePointName.Text = ""
End Sub
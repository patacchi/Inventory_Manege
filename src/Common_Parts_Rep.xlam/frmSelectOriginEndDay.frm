VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectOriginEndDay 
   Caption         =   "継承元棚卸締切日選択画面"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   OleObjectBlob   =   "frmSelectOriginEndDay.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectOriginEndDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'OirginEndDaySelect
Private Sub btnSelectComp_Click()
    'frmTanaBinのグローバル変数にセット
    frmTanaBincard.strOriginEndDay = cmbBoxOriginEndDay.Text
    Unload Me
End Sub
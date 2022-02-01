VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTanaBincard 
   Caption         =   "棚卸BINカードチェック用フォーム"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8160
   OleObjectBlob   =   "frmTanaBincard.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTanaBincard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'フォーム内共有変数
Private clsADOfrmBIN As clsADOHandle
Private clsINVDBfrmBIN As clsINVDB
Private clsEnumfrmBIN As clsEnum
Private clsSQLBc As clsSQLStringBuilder
Private objExcelFrmBIN As Excel.Application
Private dicoObjNameToFieldName As Dictionary
Private clsIncrementalfrmBIN As clsIncrementalSerch
Private Sub btnRegistTanaCSVtoDB_Click()
    '最初にCSVファイルを選択してもらう
    Dim strCSVFullPath As String
    'カレントディレクトリをダウンロードディレクトリに変更する
    Call ChCurrentDirW(GetDownloadPath)
    MsgBox "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"
    strCSVFullPath = CStr(Application.GetOpenFilename("CSVファイル,*.csv", 1, "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"))
    If strCSVFullPath = "False" Then
        'キャンセルボタンが押された
        MsgBox "キャンセルしました"
        Exit Sub
    End If
    Dim longAffected As Long
#If DontRemoveZaikoSH Then
    'DLしたファイルを残しておく（テスト環境向け）
    '取得したファイル名を引数にしてDBに登録（拡張子によって処理が分岐されるはず）
    longAffected = clsINVDBfrmBIN.UpsertINVPartsMasterfromZaikoSH(strCSVFullPath, objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc, True)
#Else
    longAffected = clsINVDBfrmBIN.UpsertINVPartsMasterfromZaikoSH(strCSVFullPath, objExcelFrmBIN, clsINVDBfrmBIN, clsADOfrmBIN, clsEnumfrmBIN, clsSQLBc, False)
#End If
End Sub
'フォーム初期化動作
Private Sub UserForm_Initialize()
    'メンバインスタンス変数セット
    If clsADOfrmBIN Is Nothing Then
        Set clsADOfrmBIN = CreateclsADOHandleInstance
    End If
    If clsINVDBfrmBIN Is Nothing Then
        Set clsINVDBfrmBIN = CreateclsINVDB
    End If
    If clsEnumfrmBIN Is Nothing Then
        Set clsEnumfrmBIN = CreateclsEnum
    End If
    If clsSQLBc Is Nothing Then
        Set clsSQLBc = CreateclsSQLStringBuilder
    End If
    If objExcelFrmBIN Is Nothing Then
        Set objExcelFrmBIN = New Excel.Application
    End If
    If dicoObjNameToFieldName Is Nothing Then
        Set dicoObjNameToFieldName = New Dictionary
    End If
    If clsIncrementalfrmBIN Is Nothing Then
        Set clsIncrementalfrmBIN = CreateclsIncrementalSerch
    End If
End Sub
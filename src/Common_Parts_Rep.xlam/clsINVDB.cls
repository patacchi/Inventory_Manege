VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsINVDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''Author Patacchi 2022_01_10
'''在庫情報DBを扱うクラス
Option Explicit
'------------------------------------------------------------------------------------------------------
'メンバ変数定義
Private strINVDBFullPath As String                                  'INV_Manege.accdbのフルパス、基本的には未指定でデフォルトの場所を読みに行く
Private strBKZaikoInfoFullPath As String                            '在庫情報のExcelファイルのフルパス、毎回ファイル名が変動するのでデフォルト値は無し
'------------------------------------------------------------------------------------------------------
'プロパティ
'INVDBFullPath
Property Get INVDBFullPath() As String
    If strINVDBFullPath <> "" Then
        'メンバ変数に既にINVDBが入っていたらそのまま返す
        INVDBFullPath = strINVDBFullPath
        Exit Property
    End If
    'メンバ変数が空だった場合、初期値を設定してやる
    Dim dbINVBDBFullpath As clsADOHandle
    Set dbINVBDBFullpath = New clsADOHandle
    Dim fsoINVDBFullPath As FileSystemObject
    Set fsoINVDBFullPath = New FileSystemObject
    'clsAdoHandleよりDBデフォルトディレクトリを、INV_CONSTよりデフォルトDBファイル名を取得し、フルパスに組み立て返す
    INVDBFullPath = fsoINVDBFullPath.BuildPath(dbINVBDBFullpath.DBPath, INV_CONST.INV_DB_FILENAME)
    Set dbINVBDBFullpath = Nothing
    Set fsoINVDBFullPath = Nothing
    Exit Property
End Property
Property Let INVDBFullPath(strargINVDBFullPath As String)
    If strargINVDBFullPath = "" Then
        '引数が空の時は抜ける
        DebugMsgWithTime "Property let INVDBFullPath:arg empty"
        Exit Property
    End If
    'メンバ変数にセットしてやる
    strINVDBFullPath = strargINVDBFullPath
    Exit Property
End Property
'BKZaikoInfoFullPath
Property Get BKZAikoInfoFullPath() As String
    '在庫情報シートのBook名に初期値は無い為、メンバ変数をそのまま返す
    BKZAikoInfoFullPath = strBKZaikoInfoFullPath
    Exit Property
End Property
Property Let BKZAikoInfoFullPath(strargBKZaikoInfoFullPath As String)
    If strargBKZaikoInfoFullPath = "" Then
        '引数が空だったら抜ける
        DebugMsgWithTime "Property let BKZaikoInfoFullPath: arg empty"
        Exit Property
    End If
    'メンバ変数にセットしてやる
    strBKZaikoInfoFullPath = strargBKZaikoInfoFullPath
    Exit Property
End Property
'------------------------------------------------------------------------------------------------------
'メソッド
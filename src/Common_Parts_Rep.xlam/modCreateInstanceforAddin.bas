Attribute VB_Name = "modCreateInstanceforAddin"
Option Explicit
'''アドイン公開向けにクラスのインスタンスを作成することを目的としたモジュール
'''公開する際にPublicNOTCreateになるため
'使用例 Dim clsAdo as clsAdoHandle
'       set clsAdo = CreateclsADOHandleInstance()
'clsAdoHandle
Public Function CreateclsADOHandleInstance() As clsADOHandle
    Dim T As clsADOHandle
    Set T = New clsADOHandle
    Set CreateclsADOHandleInstance = T
    Set T = Nothing
    Exit Function
End Function
'clsSQLStringBuilder
Public Function CreateclsSQLStringBuilder() As clsSQLStringBuilder
    Dim T As clsSQLStringBuilder
    Set T = New clsSQLStringBuilder
    Set CreateclsSQLStringBuilder = T
    Set T = Nothing
    Exit Function
End Function
'clsEnum
Public Function CreateclsEnum() As clsEnum
    Dim T As clsEnum
    Set T = New clsEnum
    Set CreateclsEnum = T
    Set T = Nothing
    Exit Function
End Function
'clsINVFB
Public Function CreateclsINVDB() As clsINVDB
    Dim T As clsINVDB
    Set T = New clsINVDB
    Set CreateclsINVDB = T
    Set T = Nothing
    Exit Function
End Function
'clsIncrementalSerch
Public Function CreateclsIncrementalSerch() As clsIncrementalSerch
    Dim T As clsIncrementalSerch
    Set T = New clsIncrementalSerch
    Set CreateclsIncrementalSerch = T
    Set T = Nothing
    Exit Function
End Function
'clsGetIE
Public Function CreateclsGetIE() As clsGetIE
    Dim T As clsGetIE
    Set T = New clsGetIE
    Set CreateclsGetIE = T
    Set T = Nothing
    Exit Function
End Function
'''SQLテストフォームを表示する
Public Sub ShowfrmSQLTest()
    frmSQLTest.Show
    Exit Sub
End Sub
'''CATDBのフィールドチェックフォームを表示する
Public Sub ShowfrmFieldChange()
    frmFieldChange.Show
    Exit Sub
End Sub
'''INV_M_Partsマスター表示フォームを表示する
Public Sub ShowfrmInvPartsMaster()
    frmINV_PartsMaster_List.Show
End Sub
'''frmTanaBincardフォームを表示する
Public Sub ShowfrmINVTanaBincard()
    frmTanaBincard.Show
End Sub
'''frmBinLabeoShow
Public Sub ShowfrmBinLabel()
    frmBinLabel.Show
End Sub
'''棚番新規登録フォーム表示(モーダル)
Public Sub ShowFrmRegistNewLocation()
    frmRegistNewLocation.Show vbModal
End Sub
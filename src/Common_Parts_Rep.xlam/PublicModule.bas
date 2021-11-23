Attribute VB_Name = "PublicModule"
Option Explicit
'参照設定
'Microsoft AciteX Data Objects 2.8 Library      %ProgramFiles(x86)%\Common Files\System\msado28.tlb
'Microsoft ADO Ext. 6.0 for DDL and Security    %ProgramFiles(x86)%\Common Files\System\msadox.dll
'Microsoft Scripting Runtime                    %SystemRoot%\SysWOW64\scrrun.dll
'Microsoft DAO 3.6 Object Library               %ProgramFiles(x86)%\Common Files\Microfost Shared\DAO\dao360.dll
'-----------------------------------------------------------------------------------------------------------------------
Public Function isUnicodePath(ByVal strCurrentPath As String) As Boolean
    'パス名にUnicodeが含まれていればTrueを返し、イベント無効にする（マクロ実行しずらいよね）
    Dim strSJIS As String           'パス名を一旦SJISに変換したもの
    Dim strReUnicode As String      'SJISに変換したパス名を再度Unicodeにしたもの
    strSJIS = StrConv(strCurrentPath, vbFromUnicode)
    strReUnicode = StrConv(strSJIS, vbUnicode)
    If strReUnicode <> strCurrentPath Then
        'うにこーどとSJIS変換して戻ってきたのが違う→Unicodeあり
        isUnicodePath = True
        Exit Function
    Else
        '同じなのでうにこーどなし
        isUnicodePath = False
        Exit Function
    End If
End Function
'-----------------------------------------------------------------------------------------------------------------------
Public Function ChCurrentDirW(ByVal DirName As String) As Boolean
    'UNICODE対応ChCurrentDir
    'SetCurrentDirectoryW（UNICODE）なので
    'StrPtrを介す必要がある・・？
    SetCurrentDirectoryW StrPtr(DirName)
End Function
'-----------------------------------------------------------------------------------------------------------
'Public Sub CreateAliasTable()
'    'エイリアステーブル作成
'    '2021_09_14 Patacchi エイリアステーブル分割
'    'HeaderとKishuNameそれぞれのエイリアステーブルに分割する
'
'    Dim strSQL As String
'    Dim dbAlias As clsSQLiteHandle
'    Dim sqlbC As clsSQLStringBuilder
'    On Error GoTo ErrorCatch
'    Set dbAlias = New clsSQLiteHandle
'    Set sqlbC = New clsSQLStringBuilder
'    'Header
'    'テーブルが存在しない場合のみ実行する
'    If Not IsTableExist(Table_AliasHeader) Then
'        strSQL = ""
'        strSQL = strSQL & strTable1_NextTable & Table_AliasHeader
'        strSQL = strSQL & strTable2_Next1stField & sqlbC.addQuote(Kishu_Header) & strTable3_TEXT & strTable_NotNull & strTable_Unique & strTable4_EndRow
'        strSQL = strSQL & sqlbC.addQuote(Kishu_Origin) & strTable3_TEXT & strTable_NotNull & strTable4_EndRow
'        strSQL = strSQL & strTable4_5_PrimaryKey & sqlbC.addQuote(Kishu_Header) & strTable4_6_EndPrimary & strTable5_EndSQL
'        dbAlias.SQL = strSQL
'        Call dbAlias.DoSQL_No_Transaction
'    End If
'    'KishuName
'    'テーブルが存在しない場合のみ実行する
'    If Not IsTableExist(Table_AliasKishu) Then
'        strSQL = ""
'        strSQL = strSQL & strTable1_NextTable & Table_AliasKishu
'        strSQL = strSQL & strTable2_Next1stField & sqlbC.addQuote(Kishu_KishuName) & strTable3_TEXT & strTable_NotNull & strTable_Unique & strTable4_EndRow
'        strSQL = strSQL & sqlbC.addQuote(Kishu_Origin) & strTable3_TEXT & strTable_NotNull & strTable4_EndRow
'        strSQL = strSQL & strTable4_5_PrimaryKey & sqlbC.addQuote(Kishu_KishuName) & strTable4_6_EndPrimary & strTable5_EndSQL
'        dbAlias.SQL = strSQL
'        Call dbAlias.DoSQL_No_Transaction
'    End If
'    GoTo CloseAndExit
'ErrorCatch:
'    If Err.Number <> 0 Then
'        MsgBox Err.Number & vbCrLf & Err.Description
'    End If
'    DebugMsgWithTime "CreateAliasTable code: " & Err.Number & "Description: " & Err.Description
'    GoTo CloseAndExit
'CloseAndExit:
'    Set dbAlias = Nothing
'    Set sqlbC = Nothing
'    Exit Sub
'End Sub
'------------------------------------------------------------------------------------------------------
Public Function getArryDimmensions(ByRef varArry As Variant) As Byte
    '配列の次元数を返す（Byteまでしか対応しないよ）
    Dim byteLocalCounter As Byte
    Dim longRows As Long
    If Not IsArray(varArry) Then
        MsgBox ("配列じゃないっぽいのが来たので中止です")
        getArryDimmensions = False
        Exit Function
    End If
    byteLocalCounter = 0
    On Error Resume Next
    Do While Err.Number = 0
        byteLocalCounter = byteLocalCounter + 1
        longRows = UBound(varArry, byteLocalCounter)
    Loop
    byteLocalCounter = byteLocalCounter - 1
    Err.Clear
    getArryDimmensions = byteLocalCounter
    Exit Function
 End Function
Public Function GetLocalTimeWithMilliSec() As String
    '現在日時をミリ秒まで付けて、フォーマット済みStringとして返す
    'ISO1806形式
    'yyyy-mm-ddTHH:MM:SS.fff
    Dim strDateWithMillisec As String
    Dim timeLocalTime As SYSTEMTIME
    Call GetLocalTime(timeLocalTime)
    strDateWithMillisec = ""
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wYear, "0000")
    strDateWithMillisec = strDateWithMillisec & "-"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wMonth, "00")
    strDateWithMillisec = strDateWithMillisec & "-"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wDay, "00")
    strDateWithMillisec = strDateWithMillisec & "T"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wHour, "00")
    strDateWithMillisec = strDateWithMillisec & ":"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wMinute, "00")
    strDateWithMillisec = strDateWithMillisec & ":"
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wSecond, "00")
    strDateWithMillisec = strDateWithMillisec & "."
    strDateWithMillisec = strDateWithMillisec & Format(timeLocalTime.wMilliseconds, "000")
    GetLocalTimeWithMilliSec = strDateWithMillisec
End Function
Public Sub OutputArrayToCSV(ByRef vararg2DimentionsDataArray As Variant, ByVal strargFilePath As String, Optional ByVal strargFileEncoding As String = "UTF-8")
    '二次元配列をCSVに吐き出す
    Dim byteDimentions As Byte
    Dim objFileStream As ADODB.Stream
    Dim longRowCounter As Long
    Dim longFieldCounter As Long
    Dim strarrField() As String
    Dim strLineBuffer As String
    On Error GoTo ErrorCatch
    byteDimentions = getArryDimmensions(vararg2DimentionsDataArray)
    If Not byteDimentions = 2 Then
        MsgBox "引数に二次元配列以外が与えられました。処理を中止します。"
        DebugMsgWithTime "OutputArrayToCSV : Not 2 Dimension Array"
        Exit Sub
    End If
    Set objFileStream = New ADODB.Stream
    With objFileStream
        'エンコード指定
        .Charset = strargFileEncoding
        '改行コード指定
        .LineSeparator = adCRLF
        .Open
        '行数ループ
        For longRowCounter = LBound(vararg2DimentionsDataArray, 1) To UBound(vararg2DimentionsDataArray, 1)
            'フィールド数ループ、ここでラインバッファを組み立てる
            'まずはstring配列にフィールド情報を入れて、Joinで連結する
            ReDim strarrField(UBound(vararg2DimentionsDataArray, 2))
            For longFieldCounter = LBound(vararg2DimentionsDataArray, 2) To UBound(vararg2DimentionsDataArray, 2)
                If IsNull(vararg2DimentionsDataArray(longRowCounter, longFieldCounter)) Then
                    'Nullの場合はNULLを入入力してやる
                    strarrField(longFieldCounter) = "NULL"
                Else
                    '通常はこっち
                    strarrField(longFieldCounter) = CStr(vararg2DimentionsDataArray(longRowCounter, longFieldCounter))
                End If
            Next longFieldCounter
            strLineBuffer = Join(strarrField, ",")
            .WriteText strLineBuffer, adWriteLine
        Next longRowCounter
        'ループが終わったらテキストファイル書き出す（上書き保存）
        .SaveToFile strargFilePath, adSaveCreateOverWrite
        .Close
    End With
    MsgBox "CSV出力完了 " & strargFilePath
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "OutputArrayToCSV code: " & Err.Number & " Description: " & Err.Description
    Exit Sub
End Sub
'''Author Daisuke Oota 2021_10_18
'''デバッグ出力時に日時も一緒に出してやる関数
'''---------------------------------------------------------------------------------------------------------------------------
Public Sub DebugMsgWithTime(strargDebugMsg As String)
    If strargDebugMsg = "" Then
        '文字列が空白だったら抜ける
        Exit Sub
    End If
    '日時込みで値を出力
    Debug.Print GetLocalTimeWithMilliSec & " " & strargDebugMsg
    Exit Sub
End Sub
'''---------------------------------------------------------------------------------------------------------------------------
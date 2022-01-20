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
'定数定義
Private Const FILTER_NAME_RANGE_PART As String = "FilterDatabase"
'------------------------------------------------------------------------------------------------------
'メンバ変数定義
Private objExcel As Excel.Application
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
'''Autor Daisuke_oota 2021_01_17
'''クラス初期化時に行う処理
Private Sub ClassInitialize()
End Sub
'''Autor Daisuke_oota 2021_01_17
'''クラス初破棄時に行う処理
Private Sub ClassFinarize()
    If Not objExcel Is Nothing Then
        'ExcelObjectがNothingじゃない場合
        objExcel.Quit
        Set objExcel = Nothing
    End If
End Sub
'------------------------------------------------------------------------------------------------------
'コンストラクタ、デストラクタ
Private Sub Class_Initialize()
    Call ClassInitialize
End Sub
Private Sub Class_Terminate()
    Call ClassFinarize
End Sub
'------------------------------------------------------------------------------------------------------
'メソッド
'''Author Patacchi 2022_01_16
'''Excel関連ファイルにオートフィルタを設定し、フィルタ範囲の名前を返す（Select の From の引数になる）
'''戻り値 string配列 1次元目はファイルのフルパス、2次元目がTable名
'''parms
'''optional strargExceFileFullpath          Excelファイルのフルパスを指定する、未指定の場合はクラスのBKZAikoInfoFullPathプロパティを設定
'''optional strargSheetName                 フィルタ設定するシート名を指定する、未指定の場合は最初のシートを対象とする
'''Optional strargKeyFieldName              フィルタ設定する際の基準フィールド名、未指定の場合手配コードとする
Public Function GetFilterRangeNameFromExcel(Optional strargExcelFileFullpath As String, Optional strargSheetName As String, Optional strargKeyFieldName As String) As String()
    'Excelファイルが指定されなかった場合はクラスのプロパティから引っ張る
    If strargExcelFileFullpath = "" Then
        strargExcelFileFullpath = Me.BKZAikoInfoFullPath
    End If
    If strargExcelFileFullpath = "" Then
        'プロパティから引っ張ってもBook名が空だった場合は抜ける
        DebugMsgWithTime "GetFilterRangeNameFromExcel: Book filename empty"
        GoTo CloseandExit
    End If
    On Error GoTo ErrorCatch
    '引数のファイルの存在確認
    Dim fsoFilterRange As FileSystemObject
    Set fsoFilterRange = New FileSystemObject
    If Not fsoFilterRange.FileExists(strargExcelFileFullpath) Then
        'ファイルが見つからなかった
        DebugMsgWithTime "GetFilterRangeNameFromExcel: file not found: " & strargExcelFileFullpath
        GoTo CloseandExit
    End If
    Dim EnumValue As clsEnum
    Set EnumValue = CreateclsEnum
    Select Case LCase(fsoFilterRange.GetExtensionName(strargExcelFileFullpath))
    Case EnumValue.DBFileExetension(xlam_dext), EnumValue.DBFileExetension(xlsm_dext), _
    EnumValue.DBFileExetension(xls_dext), EnumValue.DBFileExetension(xlsb_dext), EnumValue.DBFileExetension(xlsx_dext)
        'エクセル関連ファイルの時
        '非表示で処理するために、ApplicationオブジェクトとWorkbookオブジェクトを別に定義する
        If objExcel Is Nothing Then
            'クラスのobjExcelが初期化されていなかったらここで行う
            Set objExcel = New Excel.Application
        End If
        Dim wkbFilter As Workbook
        'workbookオブジェクトを取得
        Set wkbFilter = objExcel.Workbooks.Open(strargExcelFileFullpath)
        If wkbFilter.Worksheets.Count < 1 Then
            'シート数が1未満だった場合は抜ける
            DebugMsgWithTime "GetFilterRangeNameFromExcel: Sheet count under 1"
            GoTo CloseandExit
        End If
        If strargSheetName = "" Then
            '引数でシート名を与えられなかった場合
            '最初のシート名を対象とする
            strargSheetName = wkbFilter.Worksheets(1).Name
        End If
        '存在しないシート名を開いた場合はErr.Number = 9 、インデックス外エラーが発生するので、エラートラップを行う
        Err.Clear
        Dim shtFilter As Worksheet
        '在庫情報シートのオブジェクトを取得する、このタイミングでシートが存在しない場合はエラーが発生する
        On Error Resume Next
        Set shtFilter = wkbFilter.Worksheets(strargSheetName)
        On Error GoTo ErrorCatch
        'Err.Numberが0以外の時は処理を中断
        If Err.Number <> 0 Then
            DebugMsgWithTime "GetFilterRangeNameFromExcel: sheet not found: " & strargSheetName
            GoTo CloseandExit
            Exit Function
        End If
        Dim flgSave As Boolean
        flgSave = False
        If shtFilter.AutoFilterMode = False Then
            '在庫情報シートにフィルターが設定されていない場合
            Dim rngZaikoInfoColumn As Range
            'キーField名を決定する
            If strargKeyFieldName = "" Then
                '引数でキーフィールド名が指定されていない場合は、手配コードをデフォルトとする
                strargKeyFieldName = EnumValue.shZaikoField(F_Tehai_Code_ShZ)
            End If
            '在庫情報の列名のうち一つを検索し、Rangeオブジェクトを得る
            Set rngZaikoInfoColumn = shtFilter.Cells.Find(strargKeyFieldName)
            If Not rngZaikoInfoColumn Is Nothing Then
                '手配コードの列が見つかった場合
                '手配コードの列を基準にしてオートフィルタを設定する
                rngZaikoInfoColumn.AutoFilter
                'SaveFlagを立てる
                flgSave = True
            End If
        End If
        '名前定義すべてに対してループし、非表示の名前範囲を表示する設定にし
        'フィルター情報の名前定義を関数の戻り値として設定してやる
        Dim elmName As Name
        For Each elmName In shtFilter.Names
            If elmName.Visible = False Then
                '名前定義が非表示なっていた場合
                elmName.Visible = True
                '保存フラグを立てる
                flgSave = True
            End If
        Next elmName
        '保存フラグの状態を調べる
        If flgSave Then
            '保存フラグが立っていたらブックを保存する
            wkbFilter.Save
        End If
        'Bookへの接続を閉じてやる
        wkbFilter.Close
        Set wkbFilter = Nothing
        '処理完了したExcelファイルのテーブル一覧を取得し、フィルター識別文字列の入っているテーブル名が見つかったらそれを返す
        Dim dbGetTableList As clsADOHandle
        Set dbGetTableList = CreateclsADOHandleInstance
        dbGetTableList.DBFileName = fsoFilterRange.GetFileName(strargExcelFileFullpath)
        dbGetTableList.DBPath = fsoFilterRange.GetParentFolderName(strargExcelFileFullpath)
        Dim arrstrFilterTables() As String
        arrstrFilterTables = dbGetTableList.GetTableList
        Dim arrstrResult() As String
        arrstrResult = Filter(arrstrFilterTables, FILTER_NAME_RANGE_PART)
        If UBound(arrstrResult) = 0 Then
            '結果が一つのみであった場合（複数あると困る）
            Dim arrstrFilterRangeResult(0, 1) As String
            arrstrFilterRangeResult(0, 0) = strargExcelFileFullpath
            arrstrFilterRangeResult(0, 1) = "[" & arrstrResult(0) & "]"
            GetFilterRangeNameFromExcel = arrstrFilterRangeResult
           GoTo CloseandExit
           Exit Function
        End If
    Case EnumValue.DBFileExetension(csv_dext)
        'CSVファイルだった場合
        '一連の処理を行い、xlsxファイルにした後、そのファイル名で自信を再帰呼び出しする
        Dim strResultxlsxFileName As String
        strResultxlsxFileName = CSVFileSaveAsXlsxFileandPasteValues(strargExcelFileFullpath)
        '自身を再起呼び出し
        Dim arrstrRangeTableName() As String
        arrstrRangeTableName = Me.GetFilterRangeNameFromExcel(strResultxlsxFileName)
        '帰ってきたテーブル名を結果として返す
        GetFilterRangeNameFromExcel = arrstrRangeTableName
        GoTo CloseandExit
        Exit Function
    Case Else
        'Excelファイルじゃなかった場合
        DebugMsgWithTime "SetExcelAutoFormat: not excel file"
        GoTo CloseandExit
    End Select
ErrorCatch:
    DebugMsgWithTime "GetFilterRangeNameFromExcel code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseandExit
    Exit Function
CloseandExit:
    Set shtFilter = Nothing
    If Not wkbFilter Is Nothing Then
        'WorkbookオブジェクトがNothingじゃない場合
        wkbFilter.Close
        Set wkbFilter = Nothing
    End If
    Set dbGetTableList = Nothing
    Set EnumValue = Nothing
    Set fsoFilterRange = Nothing
    Exit Function
End Function
'''Author Patacchi 2022_01_16
'''CSVファイルをxlsxファイルに保存し、かつ全範囲値貼り付け状態にし保存したファイル名を返す
'''戻り値 string 保存後のxlsxファイル名のフルパスを返す
Private Function CSVFileSaveAsXlsxFileandPasteValues(strargCSVFileFullPath As String) As String
    On Error GoTo ErrorCatch
    If strargCSVFileFullPath = "" Then
        DebugMsgWithTime "CSVFileSaveAsxlsx: arg csv file name empty"
        GoTo CloseandExit
    End If
    'ファイル存在確認
    Dim fsoCSV As FileSystemObject
    Set fsoCSV = New FileSystemObject
    If Not fsoCSV.FileExists(strargCSVFileFullPath) Then
        DebugMsgWithTime "CSVFileSaveAsxlsx: file not found: " & strargCSVFileFullPath
        GoTo CloseandExit
        Exit Function
    End If
    '拡張子確認
    Dim EnumValue As clsEnum
    Set EnumValue = CreateclsEnum
    If Not LCase(fsoCSV.GetExtensionName(strargCSVFileFullPath)) = EnumValue.DBFileExetension(csv_dext) Then
        '拡張子がcsvじゃなかった
        DebugMsgWithTime "CSVFileSaveAsxlsx: not csv files (extention)"
        GoTo CloseandExit
        Exit Function
    End If
    'CSVファイルを開き、xlsx形式で保存する
    Dim objExcel As Excel.Application
    Set objExcel = New Excel.Application
    Dim wkbCSV As Workbook
    'workbookオブジェクトを取得
    Set wkbCSV = objExcel.Workbooks.Open(strargCSVFileFullPath)
    'とりあえずそのままxlsxとして保存する
    Dim strxlsxFileName As String
    strxlsxFileName = strargCSVFileFullPath & "." & EnumValue.DBFileExetension(xlsx_dext)
    If fsoCSV.FileExists(strxlsxFileName) Then
        '保存先のファイル名があった場合は消去する
        fsoCSV.DeleteFile strxlsxFileName, True
    End If
    wkbCSV.SaveAs strxlsxFileName, xlOpenXMLWorkbook
    wkbCSV.Close
    Set wkbCSV = Nothing
    '今度は保存したxlsxファイルに対して、全データ値貼り付けの処理を行う
    Dim wkbXlsx As Workbook
    Set wkbXlsx = objExcel.Workbooks.Open(strxlsxFileName)
    If wkbXlsx.Worksheets.Count < 1 Then
        'ワークシートの数が1未満だったら抜けるよ
        DebugMsgWithTime "CSVFileSaveAsXlsx: no contents"
        GoTo CloseandExit
        Exit Function
    End If
    Dim shXlsx As Worksheet
    'xlsxの先頭シートを取得
    Set shXlsx = wkbXlsx.Worksheets(1)
    '最終セルまでの範囲を取得し、値貼り付け
    Dim lngMaxRow As Long
    Dim lngMaxColumn As Long
    shXlsx.Range(shXlsx.Cells(1, 1), shXlsx.Cells.SpecialCells(xlCellTypeLastCell)).Copy
    DoEvents
    shXlsx.Range(shXlsx.Cells(1, 1), shXlsx.Cells.SpecialCells(xlCellTypeLastCell)).PasteSpecial Paste:=xlPasteValues
    'コピーモード解除
    objExcel.CutCopyMode = False
    '改めてxlsxファイルを保存
    wkbXlsx.Save
    wkbXlsx.Close
    Set wkbXlsx = Nothing
    CSVFileSaveAsXlsxFileandPasteValues = strxlsxFileName
    GoTo CloseandExit
ErrorCatch:
    CSVFileSaveAsXlsxFileandPasteValues = ""
    DebugMsgWithTime "CSVFileSaveAsXlsxFileandPasteValues code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseandExit
CloseandExit:
    Set shXlsx = Nothing
    If Not wkbXlsx Is Nothing Then
        wkbXlsx.Close
        Set wkbXlsx = Nothing
    End If
    If Not wkbCSV Is Nothing Then
        wkbCSV.Close
        Set wkbCSV = Nothing
    End If
    Set EnumValue = Nothing
    Set fsoCSV = Nothing
    Exit Function
End Function
'''Author Daisuke Oota 2022_01_20
'''
'''
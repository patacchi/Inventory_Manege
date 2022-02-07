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
Private Const TANA_CSV_ENDDAY_KEY As String = "棚卸締切日"
Private Const SH_FILTERD_PREFIX As String = "SHFlt"  'Sheet Filtered
'{0}    T_INV_CSV
'{1}    (afterIN word)
'{2}    F_EndDay
'{3}    (End Day String)
Private Const SQL_INPUT_ENDDAY As String = "UPDATE {0} IN""""{1} SET {2} = {3} WHERE {2} IS NULL OR {2} = """""
'------------------------------------------------------------------------------------------------------
'メンバ変数定義
Private strINVDBFullPath As String                                  'INV_Manege.accdbのフルパス、基本的には未指定でデフォルトの場所を読みに行く
Private strBKZaikoInfoFullPath As String                            '在庫情報のExcelファイルのフルパス、毎回ファイル名が変動するのでデフォルト値は無し
'インスタンス共有クラスメンバ変数
Private objExcel As Excel.Application
Private clsADOINVDB As clsADOHandle
Private clsINVDBShare As clsINVDB
Private clsEnumINVDB     As clsEnum
Private clsSQLBcINVDB As clsSQLStringBuilder
'------------------------------------------------------------------------------------------------------
'プロパティ
'INVDBFullPath
Private Property Get INVDBFullPath() As String
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
Private Property Let INVDBFullPath(strargINVDBFullPath As String)
    If strargINVDBFullPath = "" Then
        '引数が空の時は抜ける
        DebugMsgWithTime "Property let INVDBFullPath:arg empty"
        Exit Property
    End If
    'メンバ変数にセットしてやる
    strINVDBFullPath = strargINVDBFullPath
    Exit Property
End Property
'objExcel As Excel.Application
Private Property Get objExcelRef() As Excel.Application
    Set objExcelRef = objExcel
    Exit Property
End Property
Private Property Set objExcelRef(argobjExcel As Excel.Application)
    If Not argobjExcel Is Nothing Then
        Set objExcel = argobjExcel
        Exit Property
    End If
End Property
'clsADOINVDB As clsADOHandle
Private Property Get clsADORef() As clsADOHandle
    Set clsADORef = clsADOINVDB
    Exit Property
End Property
Private Property Set clsADORef(argclsADO As clsADOHandle)
    If Not argclsADO Is Nothing Then
        Set clsADOINVDB = argclsADO
        Exit Property
    End If
End Property
'Private clsINVDBShare As clsINVDB
Private Property Get clsINVDBRef() As clsINVDB
    Set clsINVDBRef = clsINVDBShare
    Exit Property
End Property
Private Property Set clsINVDBRef(argclsINVDB As clsINVDB)
    If Not argclsINVDB Is Nothing Then
        Set clsINVDBShare = argclsINVDB
        Exit Property
    End If
End Property
'Private clsEnumINVDB     As clsEnum
Private Property Get clsEnumRef() As clsEnum
    Set clsEnumRef = clsEnumINVDB
    Exit Property
End Property
Private Property Set clsEnumRef(argclsEnum As clsEnum)
    If Not argclsEnum Is Nothing Then
        Set clsEnumINVDB = argclsEnum
        Exit Property
    End If
End Property
'Private clsSQLBc As clsSQLStringBuilder
Private Property Get clsSQLBcRef() As clsSQLStringBuilder
    Set clsSQLBcRef = clsSQLBcINVDB
    Exit Property
End Property
Private Property Set clsSQLBcRef(argclsSQLBc As clsSQLStringBuilder)
    If Not argclsSQLBc Is Nothing Then
        Set clsSQLBcINVDB = argclsSQLBc
        Exit Property
    End If
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
'''クラスの共有変数を引数にとり、クラスのプロパティにセットする関数
'''Retourn Bool     引数チェックに成功したらTrue、それ以外はFalse
'''clsADORef             ADOHandleのインスタンス、繰り返し処理をすることになると思うので、このメソッド以前で宣言し、共有する
'''                         DBPath,DBFilenameは引数で渡されるクラスのインスタンスのプロパティで設定することとする
'''clsINVDBUpsert           INVDBクラスのインスタンスを指定する。このメソッド以前に作成していたら再利用する（コンストラクタの動作が重いため）
'''clsEnumRef            Enunクラスのインスタンスを指定する、できれば事前に一度だけ作成したのを使いまわしたい
'''objExcel                 objExcelのインスタンス共有のための・・・
'''NotDeleteSHFile          TrueをセットするとUpsert完了後に元データを削除しないで残す、デフォルトはFalse（消去する）
Public Function SetShareInsance(ByRef objargExcel As Excel.Application, ByRef clsargINVDB As clsINVDB, ByRef clsargADO As clsADOHandle, _
ByRef clsargEnum As clsEnum, ByRef clsargclsSQLBcRef As clsSQLStringBuilder) As Boolean
    On Error GoTo ErrorCatch
    'インスタンス変数がNothingじゃなかったら、とりあえずクラスのプロパティにセットしてやる
    'objExcel
    If Not objargExcel Is Nothing And objExcelRef Is Nothing Then
        Set objExcelRef = objargExcel
    End If
    'clsINVDBUpsert
    If Not clsargINVDB Is Nothing And clsINVDBRef Is Nothing Then
        Set clsINVDBRef = clsargINVDB
    End If
    'clsADOHandle
    If Not clsargADO Is Nothing And clsADORef Is Nothing Then
        Set clsADORef = clsargADO
'        DebugMsgWithTime "UpsertINVPartsMasterfromZaikoSH: warning! ADOUpsert instance is nothing.will delay..."
'        Set clsADORef = CreateclsADOHandleInstance
    End If
    'clsEnum
    If Not clsargEnum Is Nothing And clsEnumRef Is Nothing Then
        Set clsEnumRef = clsargEnum
    End If
    'clsclsSQLBcRef
    If Not clsargclsSQLBcRef Is Nothing And clsSQLBcRef Is Nothing Then
        Set clsSQLBcRef = clsargclsSQLBcRef
    End If
    'インスタンス変数の初期化確認
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "UpsertINVPartsMasterfromZaikoSH : Instance not Initialized"
        SetShareInsance = False
        GoTo CloseAndExit
        Exit Function
    End If
    '成功
    SetShareInsance = True
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SetShareInsance code: " & err.Number & " Description: " & err.Description
    SetShareInsance = False
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
'''インスタンス共有変数の初期化状態を調べる関数
'''戻り値   Bool    全ての変数がNothingじゃなかったら True それ以外は False
Private Function CheckInstance() As Boolean
    If objExcelRef Is Nothing Or clsADORef Is Nothing Or clsEnumRef Is Nothing Or clsINVDBRef Is Nothing Or clsSQLBcRef Is Nothing Then
        'どれか一つでもNothingだった
        CheckInstance = False
        Exit Function
    End If
    '全ての変数が初期化済み
    CheckInstance = True
    Exit Function
End Function
'''Author Patacchi 2022_01_16
'''Excel関連ファイルにオートフィルタを設定し、フィルタ範囲の名前を返す（Select の From の引数になる）
'''戻り値 string配列 1次元目はファイルのフルパス、2次元目がTable名
'''                 この時のファイルパスは、中間生成ファイルとして扱うので、オリジナルとは別に一時ファイル名を設定する
'''parms
'''optional strargExceFileFullpath          Excelファイルのフルパスを指定する、未指定の場合はクラスのBKZAikoInfoFullPathプロパティを設定
'''optional strargSheetName                 フィルタ設定するシート名を指定する、未指定の場合は最初のシートを対象とする
'''Optional strargKeyFieldName              フィルタ設定する際の基準フィールド名、未指定の場合手配コードとする
'''optional DeleteOriginFile                最後に引数で与えられたファイルを削除する場合は指定する 使いどころは、CSVで再帰呼び出しする時とか
Private Function GetFilterRangeNameFromExcel(Optional strargExcelFileFullpath As String, Optional strargSheetName As String, Optional strargKeyFieldName As String, Optional DeleteOriginFile As Boolean = False) As String()
    'インスタンス共有変数の初期化状況のチェック
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "GetFilterRangeNameFromExcel : some instance not initialized"
        GoTo CloseAndExit
        Exit Function
    End If
    'Excelファイルが指定されなかった場合はクラスのプロパティから引っ張る
    If strargExcelFileFullpath = "" Then
        strargExcelFileFullpath = Me.BKZAikoInfoFullPath
    End If
    If strargExcelFileFullpath = "" Then
        'プロパティから引っ張ってもBook名が空だった場合は抜ける
        DebugMsgWithTime "GetFilterRangeNameFromExcel: Book filename empty"
        GoTo CloseAndExit
    End If
    On Error GoTo ErrorCatch
    '引数のファイルの存在確認
    Dim fsoFilterRange As FileSystemObject
    Set fsoFilterRange = New FileSystemObject
    If Not fsoFilterRange.FileExists(strargExcelFileFullpath) Then
        'ファイルが見つからなかった
        DebugMsgWithTime "GetFilterRangeNameFromExcel: file not found: " & strargExcelFileFullpath
        GoTo CloseAndExit
    End If
    Select Case LCase(fsoFilterRange.GetExtensionName(strargExcelFileFullpath))
    Case clsEnumRef.DBFileExetension(xlam_dext), clsEnumRef.DBFileExetension(xlsm_dext), _
    clsEnumRef.DBFileExetension(xls_dext), clsEnumRef.DBFileExetension(xlsb_dext), clsEnumRef.DBFileExetension(xlsx_dext)
        'エクセル関連ファイルの時
        '非表示で処理するために、ApplicationオブジェクトとWorkbookオブジェクトを別に定義する
        Dim wkbFilter As Workbook
        'フィルタ処理後のファイルはオリジナルと別に作成し、後で中間生成ファイルとしたいので、MilliSec付きの一時ファイル名とし
        'オリジナルファイルを一時ファイル名としてコピーして、そちらで処理を行う
        Dim strOutPutFileFullPath As String
        Dim strarrNewFile(3) As String
        strarrNewFile(0) = SH_FILTERD_PREFIX
        strarrNewFile(1) = GetTimeForFileNameWithMilliSec
        strarrNewFile(2) = "."
        strarrNewFile(3) = fsoFilterRange.GetExtensionName(strargExcelFileFullpath)
        strOutPutFileFullPath = fsoFilterRange.BuildPath(fsoFilterRange.GetParentFolderName(strargExcelFileFullpath), _
        Join(strarrNewFile, ""))
        '引数で与えられたファイルを一時ファイル名にコピーする（無いとは思うけど強制上書き）
        fsoFilterRange.CopyFile strargExcelFileFullpath, strOutPutFileFullPath, True
        'ここでオリジナルファイル削除フラグが立っていたらオリジナルファイルを消去してしまう
        If DeleteOriginFile Then
            Kill strargExcelFileFullpath
        End If
        'workbookオブジェクトを取得
        Set wkbFilter = objExcelRef.Workbooks.Open(strOutPutFileFullPath)
        If wkbFilter.Worksheets.Count < 1 Then
            'シート数が1未満だった場合は抜ける
            DebugMsgWithTime "GetFilterRangeNameFromExcel: Sheet count under 1"
            GoTo CloseAndExit
        End If
        If strargSheetName = "" Then
            '引数でシート名を与えられなかった場合
            '最初のシート名を対象とする
            'この時、シート名が長すぎるとフィルタ適用テーブルが生成されないようなので、シート名を変更してしまう
            wkbFilter.Worksheets(1).Name = fsoFilterRange.GetBaseName(strOutPutFileFullPath)
            strargSheetName = wkbFilter.Worksheets(1).Name
        End If
        '存在しないシート名を開いた場合はErr.Number = 9 、インデックス外エラーが発生するので、エラートラップを行う
        err.Clear
        Dim shtFilter As Worksheet
        '在庫情報シートのオブジェクトを取得する、このタイミングでシートが存在しない場合はエラーが発生する
        On Error Resume Next
        err.Clear
        Set shtFilter = wkbFilter.Worksheets(strargSheetName)
        'Err.Numberが0以外の時は処理を中断
        If err.Number <> 0 Then
            DebugMsgWithTime "GetFilterRangeNameFromExcel: sheet not found: " & strargSheetName
            GoTo CloseAndExit
            Exit Function
        End If
        On Error GoTo ErrorCatch
        Dim flgSave As Boolean
        flgSave = False
        'フィルタが既に設定されている場合は、そのファイルに対して編集を行っても自動的にフィルタ範囲が拡大されないので、一旦解除してやる 2022_01_30 Patacchi
        If shtFilter.AutoFilterMode = True Then
            Dim rngZaikoInfoColumn As Range
            Set rngZaikoInfoColumn = shtFilter.Cells.Find(strargKeyFieldName)
            If Not rngZaikoInfoColumn Is Nothing Then
                '手配コードの列が見つかった場合
                '手配コードの列を基準にしてオートフィルタを解除する
                rngZaikoInfoColumn.AutoFilter
                'SaveFlagを立てる
                flgSave = True
            End If
        End If
        If shtFilter.AutoFilterMode = False Then
            '在庫情報シートにフィルターが設定されていない場合
            'キーField名を決定する
            If strargKeyFieldName = "" Then
                '引数でキーフィールド名が指定されていない場合は、手配コードをデフォルトとする
                strargKeyFieldName = clsEnumRef.shZaikoField(F_Tehai_Code_ShZ)
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
'        '名前定義すべてに対してループし、非表示の名前範囲を表示する設定にし
'        'フィルター情報の名前定義を関数の戻り値として設定してやる
'        Dim elmName As Name
'        For Each elmName In shtFilter.Names
'            If elmName.Visible = False Then
'                '名前定義が非表示なっていた場合
'                elmName.Visible = True
'                '保存フラグを立てる
'                flgSave = True
'            End If
'        Next elmName
        '保存フラグの状態を調べる
        If flgSave Then
            '保存フラグが立っていたらブックを保存する
'            wkbFilter.SaveAs Filename:=strOutPutFileFullPath
            wkbFilter.Save
        End If
        'Bookへの接続を閉じてやる
        wkbFilter.Close
        Set wkbFilter = Nothing
        '処理完了したExcelファイルのテーブル一覧を取得し、フィルター識別文字列の入っているテーブル名が見つかったらそれを返す
        clsADORef.DBFileName = fsoFilterRange.GetFileName(strOutPutFileFullPath)
        clsADORef.DBPath = fsoFilterRange.GetParentFolderName(strOutPutFileFullPath)
        Dim arrstrFilterTables() As String
        arrstrFilterTables = clsADORef.GetTableList
        Dim arrstrResult() As String
        '帰ってきたテーブルが2個以上の場合のみフィルタ処理を行う
        '1個だけの時はそのまま採用する
        If UBound(arrstrFilterTables) - LBound(arrstrFilterTables) = 0 Then
            '結果が1個しかなかった場合
            'そのまま返す
            'この状態だとフィルタ適用後のテーブル名が無いので、失敗する可能性があるので警告を出す
            MsgBox "ファイル名が長いため、フィルタ処理がスキップされました。正常に動作しないかもしれません。" & vbCrLf & strargExcelFileFullpath
            ReDim arrstrResult(0)
            arrstrResult(0) = arrstrFilterTables(0)
        Else
            '複数テーブル返ってきたときはフィルタ処理を行う
            arrstrResult = Filter(arrstrFilterTables, FILTER_NAME_RANGE_PART)
        End If
        If UBound(arrstrResult) = 0 Then
            '結果が一つのみであった場合（複数あると困る）
            Dim arrstrFilterRangeResult(0, 1) As String
            'ここで返されるファイルは中間生成ファイルなので、後で消去したいのでファイル名は一時ファイル名とする
            arrstrFilterRangeResult(0, 0) = strOutPutFileFullPath
'            arrstrFilterRangeResult(0, 0) = strargExcelFileFullpath
            arrstrFilterRangeResult(0, 1) = "[" & arrstrResult(0) & "]"
            GetFilterRangeNameFromExcel = arrstrFilterRangeResult
           GoTo CloseAndExit
           Exit Function
        End If
    Case clsEnumRef.DBFileExetension(csv_dext)
        'CSVファイルだった場合
        '一連の処理を行い、xlsxファイルにした後、そのファイル名で自信を再帰呼び出しする
        Dim strResultxlsxFileName As String
        strResultxlsxFileName = CSVFileSaveAsXlsxFileandPasteValues(strargExcelFileFullpath)
        '自身を再起呼び出し、この時変換後のファイルは不要なので、オリジナルファイル削除フラグ付きで実行する
        Dim arrstrRangeTableName() As String
        arrstrRangeTableName = GetFilterRangeNameFromExcel(strResultxlsxFileName, , , True)
        '帰ってきたテーブル名を結果として返す
        GetFilterRangeNameFromExcel = arrstrRangeTableName
        GoTo CloseAndExit
        Exit Function
    Case Else
        'Excelファイルじゃなかった場合
        DebugMsgWithTime "SetExcelAutoFormat: not excel file"
        GoTo CloseAndExit
    End Select
ErrorCatch:
    DebugMsgWithTime "GetFilterRangeNameFromExcel code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Set shtFilter = Nothing
    If Not wkbFilter Is Nothing Then
        'WorkbookオブジェクトがNothingじゃない場合
        wkbFilter.Close
        Set wkbFilter = Nothing
    End If
    Set fsoFilterRange = Nothing
    Exit Function
End Function
'''Author Patacchi 2022_01_16
'''CSVファイルをxlsxファイルに保存し、かつ全範囲値貼り付け状態にし保存したファイル名を返す
'''戻り値 string 保存後のxlsxファイル名のフルパスを返す
Private Function CSVFileSaveAsXlsxFileandPasteValues(strargCSVFileFullPath As String) As String
    'インスタンス変数初期化チェック
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "CSVFileSaveAsXlsxFileandPasteValues : Share Instance not initialized"
        GoTo CloseAndExit
        Exit Function
    End If
    On Error GoTo ErrorCatch
    If strargCSVFileFullPath = "" Then
        DebugMsgWithTime "CSVFileSaveAsxlsx: arg csv file name empty"
        GoTo CloseAndExit
    End If
    'ファイル存在確認
    Dim fsoCSV As FileSystemObject
    Set fsoCSV = New FileSystemObject
    If Not fsoCSV.FileExists(strargCSVFileFullPath) Then
        DebugMsgWithTime "CSVFileSaveAsxlsx: file not found: " & strargCSVFileFullPath
        GoTo CloseAndExit
        Exit Function
    End If
    '拡張子確認
    If Not LCase(fsoCSV.GetExtensionName(strargCSVFileFullPath)) = clsEnumRef.DBFileExetension(csv_dext) Then
        '拡張子がcsvじゃなかった
        DebugMsgWithTime "CSVFileSaveAsxlsx: not csv files (extention)"
        GoTo CloseAndExit
        Exit Function
    End If
    'CSVファイルを開き、xlsx形式で保存する
    Dim wkbCSV As Workbook
    'workbookオブジェクトを取得
    Set wkbCSV = objExcelRef.Workbooks.Open(strargCSVFileFullPath)
    'とりあえずそのままxlsxとして保存する
    Dim strxlsxFileName As String
    strxlsxFileName = strargCSVFileFullPath & "." & clsEnumRef.DBFileExetension(xlsx_dext)
    If fsoCSV.FileExists(strxlsxFileName) Then
        '保存先のファイル名があった場合は消去する
        fsoCSV.DeleteFile strxlsxFileName, True
    End If
    wkbCSV.SaveAs strxlsxFileName, xlOpenXMLWorkbook
    wkbCSV.Close
    Set wkbCSV = Nothing
    '今度は保存したxlsxファイルに対して、全データ値貼り付けの処理を行う
    Dim wkbXlsx As Workbook
    Set wkbXlsx = objExcelRef.Workbooks.Open(strxlsxFileName)
    If wkbXlsx.Worksheets.Count < 1 Then
        'ワークシートの数が1未満だったら抜けるよ
        DebugMsgWithTime "CSVFileSaveAsXlsx: no contents"
        GoTo CloseAndExit
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
    objExcelRef.CutCopyMode = False
    '改めてxlsxファイルを保存
    wkbXlsx.Save
    wkbXlsx.Close
    Set wkbXlsx = Nothing
    CSVFileSaveAsXlsxFileandPasteValues = strxlsxFileName
    GoTo CloseAndExit
ErrorCatch:
    CSVFileSaveAsXlsxFileandPasteValues = ""
    DebugMsgWithTime "CSVFileSaveAsXlsxFileandPasteValues code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set shXlsx = Nothing
    If Not wkbXlsx Is Nothing Then
        wkbXlsx.Close
        Set wkbXlsx = Nothing
    End If
    If Not wkbCSV Is Nothing Then
        wkbCSV.Close
        Set wkbCSV = Nothing
    End If
    Set fsoCSV = Nothing
    Exit Function
End Function
'''Author Daisuke Oota 2022_01_20
'''在庫情報シートのフルパスを引数として、INVパーツマスターDBに登録するプロシージャ
'''棚卸CSVファイルをT_INV_CSVにする機能も一緒にした、拡張子で処理を分岐する
'''戻り値 long      変更があったレコード数を返す。0以下だと何かしらがあった？
'''args
'''strargZaikoSHFullPath    在庫情報シートのフルパス
'''clsADORef             ADOHandleのインスタンス、繰り返し処理をすることになると思うので、このメソッド以前で宣言し、共有する
'''                         DBPath,DBFilenameは引数で渡されるクラスのインスタンスのプロパティで設定することとする
'''clsINVDBUpsert           INVDBクラスのインスタンスを指定する。このメソッド以前に作成していたら再利用する（コンストラクタの動作が重いため）
'''clsEnumRef            Enunクラスのインスタンスを指定する、できれば事前に一度だけ作成したのを使いまわしたい
'''objExcel                 objExcelのインスタンス共有のための・・・
'''NotDeleteSHFile          TrueをセットするとUpsert完了後に元データを削除しないで残す、デフォルトはFalse（消去する）
Public Function UpsertINVPartsMasterfromZaikoSH(strargZaikoSHFullPath As String, _
ByRef objargExcel As Excel.Application, ByRef clsargINVDB As clsINVDB, ByRef clsargADO As clsADOHandle, ByRef clsargEnum As clsEnum, ByRef clsargclsSQLBcRef As clsSQLStringBuilder, _
Optional NotDeleteSHFile As Boolean = False) As Long
    On Error GoTo ErrorCatch
    'インスタンス変数がNothingじゃなかったら、とりあえずクラスのプロパティにセットしてやる
    'objExcel
    If Not objargExcel Is Nothing And objExcelRef Is Nothing Then
        Set objExcelRef = objargExcel
    End If
    'clsINVDBUpsert
    If Not clsargINVDB Is Nothing And clsINVDBRef Is Nothing Then
        Set clsINVDBRef = clsargINVDB
    End If
    'clsADOHandle
    If Not clsargADO Is Nothing And clsADORef Is Nothing Then
        Set clsADORef = clsargADO
'        DebugMsgWithTime "UpsertINVPartsMasterfromZaikoSH: warning! ADOUpsert instance is nothing.will delay..."
'        Set clsADORef = CreateclsADOHandleInstance
    End If
    'clsEnum
    If Not clsargEnum Is Nothing And clsEnumRef Is Nothing Then
        Set clsEnumRef = clsargEnum
    End If
    'clsclsSQLBcRef
    If Not clsargclsSQLBcRef Is Nothing And clsSQLBcRef Is Nothing Then
        Set clsSQLBcRef = clsargclsSQLBcRef
    End If
    'インスタンス変数の初期化確認
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "UpsertINVPartsMasterfromZaikoSH : Instance not Initialized"
        GoTo CloseAndExit
        Exit Function
    End If
    'ファイルの存在確認
    Dim fsoUpsertPartsMaster As FileSystemObject
    Set fsoUpsertPartsMaster = New FileSystemObject
    If Not fsoUpsertPartsMaster.FileExists(strargZaikoSHFullPath) Then
        '引数で指定したファイルが存在しなかった場合
        DebugMsgWithTime "UpsertINVPartsMaster file not found: " & strargZaikoSHFullPath
        UpsertINVPartsMasterfromZaikoSH = 0
        GoTo CloseAndExit
    End If
    'この時、拡張子がcsvだったらT_INV_CSV登録メソッドへ分岐する（帰ってこない）
    If fsoUpsertPartsMaster.GetExtensionName(strargZaikoSHFullPath) = clsEnumRef.DBFileExetension(csv_dext) Then
        '拡張子がCSVだった
        MsgBox TanaCSVFileToT_INV_CSV(strargZaikoSHFullPath, NotDeleteSHFile)
        GoTo CloseAndExit
    End If
    Dim strFilterdZaikoSh() As String
    'フィルタ処理後のファイル名を受け取る。この時、シート名を明示的にINT_CONSTから引っ張り指定してやる
    strFilterdZaikoSh = GetFilterRangeNameFromExcel(strargZaikoSHFullPath, INV_CONST.INV_SH_ZAIKO_NAME)
    '処理後に初期化されていない配列が帰ってきたら、在庫情報シートが存在しない可能性が高いので処理を中断する
    If Not IsRedim(strFilterdZaikoSh) Then
        DebugMsgWithTime "UpserINVPartsMaster: not found zaiko_info sheets origin file: " & strargZaikoSHFullPath
        UpsertINVPartsMasterfromZaikoSH = 0
        GoTo CloseAndExit
    End If
    'ここまでで在庫情報シートが確定したので、在庫情報からDBに登録するメソッドで処理を行う
    Dim longResultAffected As Long
    '棚卸CSVから在庫情報シートを取得した場合、AdoHandleのDBPathが変更されているので、明示的にデフォルトに戻す
    clsADORef.SetDBPathandFilenameDefault
    'まずは在庫情報シートをTempTableに格納する（直接Excelファイル参照だとうまくいかなかった）
    longResultAffected = IncludeZaikoSH_to_TempTable(strFilterdZaikoSh(0, 0), strFilterdZaikoSh(0, 1))
    UpsertINVPartsMasterfromZaikoSH = longResultAffected
    If longResultAffected < 1 Then
        '件数が1件未満だった場合は抜ける
        DebugMsgWithTime "UpsertINVPartMasterFromZaikoSH: no entry found"
        GoTo CloseAndExit
        Exit Function
    End If
    '棚番マスターに登録する際に、SELECT DISTINCT ロケーション しないと重複してる時にそのまま登録してしまう、が
    'DISTINCT等の集計を行ったデータを元にすると、更新可能なクエリであることが・・・と出てしまうので
    'DISTINCTの結果を一旦テーブルに吐き出さないとダメらしい・・・なんてこと
    Dim isCollect As Boolean
    isCollect = SELECT_ToTempTable(INV_CONST.F_SH_ZAIKO_TANA_TEXT)
    If Not isCollect Then
        MsgBox "UpsertINVPartsMasterfromZaikoSH で Select結果を一時テーブルに格納時にエラーが発生しました"
        DebugMsgWithTime "UpsertINVPartsMasterFromZaikoSH: Select to Temp Table fail"
        UpsertINVPartsMasterfromZaikoSH = 0
        Exit Function
    End If
    'TempTableから本番テーブルに登録する
    'まずは棚番マスターから
    'DBPathを初期値に戻す
    clsADORef.SetDBPathandFilenameDefault
    'SQL作成
    '置換用dic
    Dim dicReplaceParam As Dictionary
    Set dicReplaceParam = New Dictionary
    dicReplaceParam.Add 0, fsoUpsertPartsMaster.BuildPath(clsADORef.DBPath, PublicConst.TEMP_DB_FILENAME)
    dicReplaceParam.Add 1, clsSQLBcRef.addQuote(GetLocalTimeWithMilliSec)
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(SQL_INV_TEMP_TO_M_TANA_0INVDBFullPath_1LocalTimeMillisec, dicReplaceParam)
    'SQL実行前にConnectModeのWriteフラグを調査する
    If Not clsADORef.ConnectMode And adModeWrite Then
        clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    End If
    Call clsADORef.Do_SQL_with_NO_Transaction
    '次にPartsMasterを更新する
    '各種結合、抽出条件文字列作成IN句も作るの・・？
    'IN DB_Temp
    Dim strINWord As String
    strINWord = clsSQLBcRef.CreateAfterIN_WordFromSHFullPath(fsoUpsertPartsMaster.BuildPath(clsADORef.DBPath, PublicConst.TEMP_DB_FILENAME), clsEnumRef)
    'ON Tana and TTmp
    Dim strON_TanaAndTTemp
    strON_TanaAndTTemp = clsSQLBcRef.ReturnCommonFieldCondition(clsEnumRef.GetKeyCommonFieldTanaAndZaikoSH, clsEnumRef.Dic_ReverseSHZaiko, clsEnumRef.Dic_NeedTrim_ZaikoSH, INVDB_Tana_Alias_sia, INVDB_Tmp_Alias_sia, Equal_sfc, Connect_Comma_sfc, clsEnumRef)
    'ON Parts and TTmp
    Dim strON_PartsAndTTmp
    strON_PartsAndTTmp = clsSQLBcRef.ReturnCommonFieldCondition(clsEnumRef.GetKeyCommonFieldPatsAndZaikoSH, clsEnumRef.Dic_ReverseSHZaiko, clsEnumRef.Dic_NeedTrim_ZaikoSH, INVDB_Parts_Alias_sia, INVDB_Tmp_Alias_sia, Equal_sfc, Connect_Comma_sfc, clsEnumRef)
    'Set Condition
    Dim strSETExp As String
    strSETExp = clsSQLBcRef.ReturnCommonFieldCondition(clsEnumRef.GetUpdateCommonFieldZaikSHandMParts, clsEnumRef.Dic_ReverseSHZaiko, clsEnumRef.Dic_NeedTrim_ZaikoSH, INVDB_Parts_Alias_sia, INVDB_Tmp_Alias_sia, Equal_sfc, Connect_Comma_sfc, clsEnumRef)
    'Where Condition
    Dim strWhereExp As String
    strWhereExp = clsSQLBcRef.ReturnCommonFieldCondition(clsEnumRef.GetUpdateCommonFieldZaikSHandMParts, clsEnumRef.Dic_ReverseSHZaiko, clsEnumRef.Dic_NeedTrim_ZaikoSH, INVDB_Parts_Alias_sia, INVDB_Tmp_Alias_sia, NOT_Equal_sfc, Connect_OR_sfc, clsEnumRef)
    If dicReplaceParam Is Nothing Then
        '置換用DictionaryがNothingだったら初期化
        Set dicReplaceParam = New Dictionary
    End If
    '置換用Dic準備（全消去）
''T_INV_M_Parts                      {0}
''TDBPrts                            {1}
''T_INV_M_Tana                       {2}
''TDBTana                            {3}
''T_INV_Temp                         {4}
''TTmp                               {5}
''(CreateAfterINWord(DB_Temp.accdb)  {6}
''(ON Condition Tana and TTmp)       {7}
''(ON Condition Parts and TTmp)      {8}
''F_INV_Tana_ID                      {9}
''(SET Condition Parts and TTmp)     {10}
''(WHERE condition Parts ad TTmp)    {11}
''F_INV_Tehai_Code                   {12}
'{13} InputDate
'{14} (GetLocaltimeWithMilliSec)
'Public Const SQL_INV_UPSERT_PARSTABL_FROM_TTMP_AND_TANA As String = "UPDATE  {0} AS {1} " & vbCrLf &
    dicReplaceParam.RemoveAll
    dicReplaceParam.Add 0, INV_CONST.T_INV_M_Parts
    dicReplaceParam.Add 1, clsEnumRef.SQL_INV_Alias(INVDB_Parts_Alias_sia)
    dicReplaceParam.Add 2, INV_CONST.T_INV_M_Tana
    dicReplaceParam.Add 3, clsEnumRef.SQL_INV_Alias(INVDB_Tana_Alias_sia)
    dicReplaceParam.Add 4, INV_CONST.T_INV_TEMP
    dicReplaceParam.Add 5, clsEnumRef.SQL_INV_Alias(INVDB_Tmp_Alias_sia)
    dicReplaceParam.Add 6, strINWord
    dicReplaceParam.Add 7, strON_TanaAndTTemp
    dicReplaceParam.Add 8, strON_PartsAndTTmp
    dicReplaceParam.Add 9, clsEnumRef.INVMasterTana(F_INV_TANA_ID_IMT)
    dicReplaceParam.Add 10, strSETExp
    dicReplaceParam.Add 11, strWhereExp
    dicReplaceParam.Add 12, clsEnumRef.INVMasterParts(F_Tehai_Code_IMPrt)
    dicReplaceParam.Add 13, PublicConst.INPUT_DATE
    dicReplaceParam.Add 14, GetLocalTimeWithMilliSec
    'AppendFormatしてSQL作成
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(SQL_INV_UPSERT_PARSTABL_FROM_TTMP_AND_TANA, dicReplaceParam)
    'SQL実行前にConnectModeでWriteフラグが立っているか確認する
    If Not clsADORef.ConnectMode And adModeWrite Then
        'Writeフラグが立っていなかったので、立てる
        clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    End If
    'SQL実行
    isCollect = clsADORef.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        DebugMsgWithTime "UpsertINVPartsMasterFromZaikoSH: SQL Execute fail"
        MsgBox "UpsertINVPartsMasterfromZaikoSHでT_M_Parts更新時にSQL実行エラーが発生しました"
        UpsertINVPartsMasterfromZaikoSH = 0
        GoTo CloseAndExit
        Exit Function
    End If
    'Upsert成功
    If Not NotDeleteSHFile Then
        '削除回避フラグが立っていない時
        'DLした在庫情報ファイル(xls)を削除する
        Kill strargZaikoSHFullPath
    End If
    '処理件数を返して終了
    DebugMsgWithTime "Upsert Affected: " & clsADORef.Affected & " OriginFile: " & strargZaikoSHFullPath
    UpsertINVPartsMasterfromZaikoSH = clsADORef.Affected
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "UpsertINBPartMasterFromZaikoSH code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set fsoUpsertPartsMaster = Nothing
    Set dicReplaceParam = Nothing
    'INVDBインスタンスはByRefで他の物を再利用しているので解放禁止
'    If Not INVDBupsertINVPartMaster Is Nothing Then
'        'INVDB変数がNothingだった場合
'        Set INVDBupsertINVPartMaster = Nothing
'    End If
    Exit Function
End Function
'''棚卸CSVファイルからDBに登録するメソッド
'''Return   string  登録件数を文字列として返す ex)棚卸CSVよりDBに tt日分として xx 件追加しました
'''args
'''strargTanaCSVFullPath    棚卸CSVのフルパス
'''NoFileDelete             Trueにセットすると引数のファイルを消去せずに残す、デフォルトはFalse（消去する）
Private Function TanaCSVFileToT_INV_CSV(strargTanaCSVFullPath As String, Optional NoFileDelete As Boolean = False) As String
    On Error GoTo ErrorCatch
    'インスタンス変数の初期化状況確認
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "TanaCSVFileToT_INV_CSV : instance not Initialized"
        TanaCSVFileToT_INV_CSV = "インスタンス初期化チェックエラーです。"
        GoTo CloseAndExit
        Exit Function
    End If
    'CSVファイルにフィルタ処理し、ファイル名を受け取る
    Dim strFilteredSHFile() As String
    strFilteredSHFile = GetFilterRangeNameFromExcel(strargTanaCSVFullPath)
    '戻り値の配列が初期化されていなかったら失敗してるので抜ける
    If Not IsRedim(strFilteredSHFile) Then
        DebugMsgWithTime "TanaCSVFileToT_INV_CSV : Filter Range fail"
        TanaCSVFileToT_INV_CSV = "CSVファイルを変換する際にエラーが発生しました。処理を中断します"
        GoTo CloseAndExit
    End If
    'フィルタ処理後CSVファイルを一時テーブルに流し込む
    Dim longAffected As Long
    longAffected = IncludeZaikoSH_to_TempTable(strFilteredSHFile(0, 0), strFilteredSHFile(0, 1))
    If longAffected < 1 Then
        '在庫情報シートから一時テーブルに登録した件数が1件未満だった
        DebugMsgWithTime "TanaCSVFileToT_INV_CSV: Tana CSV have no data found"
        TanaCSVFileToT_INV_CSV = "棚卸CSVからのデータ取得に失敗しました"
        GoTo CloseAndExit
    End If
    'SQL組立のために必要な文字列を取得する
    'ON
    Dim strONcondition As String
    '棚卸CSVの締切日を持ってくる
    Dim strTanaEndDay As String
    strTanaEndDay = GetEndDayFromTanaCSV(strargTanaCSVFullPath)
    '締切日取得出来たらTempテーブルにフィールド追加してやり、値をセットしてしまう
    If Not AddEndDayFieldAndInputDate(strTanaEndDay) Then
        'フィールド追加失敗
        DebugMsgWithTime "TanaCSVFileToINV_CSV : fail append and select INTO Tmp table"
        TanaCSVFileToT_INV_CSV = "CSVファイルから一時テーブルに登録する際にエラーが発生しました"
        clsADORef.CloseClassConnection
        GoTo CloseAndExit
    End If
    'AfterIn
    'DBPathを標準に戻す
    clsADORef.SetDBPathandFilenameDefault
    Dim fsoTanaCSV As FileSystemObject
    Set fsoTanaCSV = New FileSystemObject
    Dim strAfterIN As String
    strAfterIN = clsSQLBcRef.CreateAfterIN_WordFromSHFullPath(fsoTanaCSV.BuildPath(clsADORef.DBPath, PublicConst.TEMP_DB_FILENAME), clsEnumRef)
    'SET
    Dim strSetCondition As String
    strSetCondition = clsSQLBcRef.ReturnCommonFieldCondition(clsEnumRef.GetCommonFieldArrayT_INV_CSV_andtemp, clsEnumRef.Dic_Reverse_INV_CSV, clsEnumRef.Dic_NeedTrim_CSV, _
    TanaCSV_Alias_sia, INVDB_Tmp_Alias_sia, Equal_sfc, Connect_Comma_sfc, clsEnumRef)
''{0}    T_INV_CSV
''{1}    TCSVTana
''{2}    T_INV_Temp
''{3}    (After IN Word) Temp_DB
''{4]    Ttmp
''{5}    手配コード
''{6]    (SET condition)
''{7}    棚卸締切日
'Public Const SQL_INV_TMP_TO_CSVTANA As String = "UPDATE {0} AS {1} " & vbCrLf & _
    '置換用Dictionary定義、構築
    Dim dicReplaceINVCSV As Dictionary
    Set dicReplaceINVCSV = New Dictionary
    dicReplaceINVCSV.Add 0, INV_CONST.T_INV_CSV
    dicReplaceINVCSV.Add 1, clsEnumRef.SQL_INV_Alias(TanaCSV_Alias_sia)
    dicReplaceINVCSV.Add 2, INV_CONST.T_INV_TEMP
    dicReplaceINVCSV.Add 3, strAfterIN
    dicReplaceINVCSV.Add 4, clsEnumRef.SQL_INV_Alias(INVDB_Tmp_Alias_sia)
    dicReplaceINVCSV.Add 5, clsEnumRef.CSVTanafield(F_Tehai_Code_ICS)
    dicReplaceINVCSV.Add 6, strSetCondition
    dicReplaceINVCSV.Add 7, clsEnumRef.CSVTanafield(F_EndDay_ICS)
    'Replace実行し、SQLを得る
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(INV_CONST.SQL_INV_TMP_TO_CSVTANA, dicReplaceINVCSV)
    'SQL実行
    Dim isCollect As Boolean
    '一旦明示的に接続を切断してやる
    clsADORef.CloseClassConnection
    'Writeフラグ立てる
    clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    isCollect = clsADORef.Do_SQL_with_NO_Transaction
    'Writeフラグを下げる
    clsADORef.ConnectMode = clsADORef.ConnectMode And Not adModeWrite
    If Not isCollect Then
        DebugMsgWithTime "TanaCSVFileToT_INV_CSV : Do sql fail... "
        TanaCSVFileToT_INV_CSV = "TanaCSVFileToT_INV_CSVでSQL実行時にエラーが発生しました。"
        GoTo CloseAndExit
    End If
    '中間生成ファイルを削除する
    DebugMsgWithTime "TanaCSVFileToT_INV_CSV : delete work file: " & strFilteredSHFile(0, 0)
    Kill strFilteredSHFile(0, 0)
    TanaCSVFileToT_INV_CSV = "棚卸CSVから 棚卸締切日： " & strTanaEndDay & " 分として " & clsADORef.Affected & " 件のデータを追加しました"
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "TanaCSVFileToT_INV_CSV code : " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set fsoTanaCSV = Nothing
    Set dicReplaceINVCSV = Nothing
    Exit Function
End Function
'''棚卸CSVから締切日を持ってくる
'''Return String    締切日をテキスト形式で
Private Function GetEndDayFromTanaCSV(TanaCSVFullPath As String) As String
    On Error GoTo ErrorCatch
    Dim wkbEndDay As Workbook
    Set wkbEndDay = objExcelRef.Workbooks.Open(TanaCSVFullPath)
    Dim ShEndDay As Worksheet
    '開いたブックの1番目（最初）のシートで決め打ちする
    Set ShEndDay = wkbEndDay.Worksheets(1)
    Dim rngEndDay As Range
    '棚卸締切日でセル検索してRangeオブジェクト取得
    Set rngEndDay = ShEndDay.Cells.Find(TANA_CSV_ENDDAY_KEY)
    '棚卸締切日のセルの右10文字を返して終了
    GetEndDayFromTanaCSV = Right(CStr(rngEndDay.Value), 10)
    GoTo CloseAndExit
ErrorCatch:
    GetEndDayFromTanaCSV = ""
    GoTo CloseAndExit
CloseAndExit:
    Set rngEndDay = Nothing
    Set ShEndDay = Nothing
    If Not wkbEndDay Is Nothing Then
        wkbEndDay.Close
        Set wkbEndDay = Nothing
    End If
    Exit Function
End Function
'''Tempテーブルに棚卸締切日フィールドを追加し、引数で指定された締切日を入力してしまうメソッド
'''Return Bool      成功したらTrueを返す、それ以外はFalse
'''args
'''strargEndDay     締切日を10桁のstring形式で
Private Function AddEndDayFieldAndInputDate(strargEndDay As String) As Boolean
    On Error GoTo ErrorCatch
    '引数チェック
    If Len(strargEndDay) > 10 Then
        '10文字超えてたら右側10文字だけ採用する
        strargEndDay = Right(strargEndDay, 10)
    End If
    'まずはTempテーブルにフィールドを追加してやる
    'DBpathをデフォルトにする
    clsADORef.SetDBPathandFilenameDefault
    'DBFilename に 一時DBファイル名を指定する
    clsADORef.DBFileName = PublicConst.TEMP_DB_FILENAME
    Dim isCollect As Boolean
    'Writeフラグを立てる
    clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    'フィールド追加実行
    isCollect = clsADORef.AppendField(INV_CONST.T_INV_TEMP, clsEnumRef.CSVTanafield(F_EndDay_ICS), Char_ShortText_typ, "(10)", PublicConst.TEMP_DB_FILENAME)
    If Not isCollect Then
        DebugMsgWithTime "AddEndDayFieldAndInputDate : fail append Field."
        GoTo CloseAndExit
        Exit Function
    End If
    '次に棚卸締切日に値をセットしてしまう
'{0}    T_INV_CSV
''{1}    (afterIN word)
''{2}    F_EndDay
''{3}    (End Day String)
'Private Const SQL_INPUT_ENDDAY As String = "UPDATE {0} IN""{1} SET {2} = {3} WHERE {2} IS NULL"
    '置換用Dic準備
    Dim dicReplaceAppend As Dictionary
    Set dicReplaceAppend = New Dictionary
    Dim fsoReplaceAppend As FileSystemObject
    Set fsoReplaceAppend = New FileSystemObject
    dicReplaceAppend.Add 0, INV_CONST.T_INV_TEMP
    dicReplaceAppend.Add 1, clsSQLBcRef.CreateAfterIN_WordFromSHFullPath(fsoReplaceAppend.BuildPath(clsADORef.DBPath, clsADORef.DBFileName), clsEnumRef)
    dicReplaceAppend.Add 2, clsEnumINVDB.CSVTanafield(F_EndDay_ICS)
    '日付はクォートしないと日付シリアルに変換しようとしてしまう
    dicReplaceAppend.Add 3, clsSQLBcRef.addQuote(strargEndDay)
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(SQL_INPUT_ENDDAY, dicReplaceAppend)
    '一旦接続を閉じる
    clsADORef.CloseClassConnection
    'Writeフラグを立てる
    clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    'SQL実行
    isCollect = clsADORef.Do_SQL_with_NO_Transaction
    'Writeフラグ下げる
    clsADORef.ConnectMode = clsADORef.ConnectMode And Not adModeWrite
    If Not isCollect Then
        DebugMsgWithTime "AddEndDayFieldAndInputDate : fail do sql"
        AddEndDayFieldAndInputDate = False
        GoTo CloseAndExit
    End If
    AddEndDayFieldAndInputDate = True
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "AddEndDayFieldAndInputDate code: " & err.Number & " Description: " & err.Description
    AddEndDayFieldAndInputDate = False
    GoTo CloseAndExit
CloseAndExit:
    'Writeフラグを下げる
    clsADORef.ConnectMode = clsADORef.ConnectMode And Not adModeWrite
    Exit Function
End Function
'''Author Daisuke oota 2022_01_21
'''在庫情報シートのフルパスを引数として、DBの一時テーブルに登録するメソッド
'''args
'''strargFilterdZaikoSHFullPath         フィルタ処理済みのブック名のフルパス
'''strargTableName                      在庫情報シートのSelect 文 Fromに入るテーブル名
'''戻り値 long      影響のあった行数を返す
Private Function IncludeZaikoSH_to_TempTable(strargFilterdZaikoSHFullPath As String, strargTableName As String) As Long
    On Error GoTo ErrorCatch
    'インスタンス変数の初期化状況の確認
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "IncludeZaikoSH_to_TempTable : instance not Initialized"
        GoTo CloseAndExit
        Exit Function
    End If
    'DBPathをデフォルトに戻す、一時ファイル格納DBも保存場所はデフォルトPathのため
    clsADORef.SetDBPathandFilenameDefault
    'このメソッドでは一時テーブルDBを扱うので、明示的にファイルを指定してやる
    clsADORef.DBFileName = PublicConst.TEMP_DB_FILENAME
    'ファイル、テーブルの空データチェック
    If strargFilterdZaikoSHFullPath = "" Or strargTableName = "" Then
        DebugMsgWithTime "IncludeZaikoSH_to_TempTable: Zaiko sheet filename or table name is empty"
        IncludeZaikoSH_to_TempTable = 0
        GoTo CloseAndExit
        Exit Function
    End If
    '在庫情報シートの内容を一時テーブルに格納する
    'まずは最初に一時テーブルを消去してやる
    Dim isCollect As Boolean
    isCollect = clsADORef.DropTable(INV_CONST.T_INV_TEMP, PublicConst.TEMP_DB_FILENAME)
    If Not isCollect Then
        DebugMsgWithTime "IncludeZaikoSH_to_TempTable: Do sql fail..."
        MsgBox "IncludeZaikoSH_to_TempTable でSQLの実行に失敗しました。" & err.Description
        GoTo CloseAndExit
    End If
    'SQL文の組み立て
    'INV_CONST.SQL_INV_SH_TO_DB_TEMPTABLE_0Table_1INword
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    dicReplaceParm.Add 0, strargTableName
    dicReplaceParm.Add 1, clsSQLBcRef.CreateAfterIN_WordFromSHFullPath(strargFilterdZaikoSHFullPath, clsEnumRef)
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(INV_CONST.SQL_INV_SH_TO_DB_TEMPTABLE_0Table_1INword, dicReplaceParm)
    'SQL実行前にConnectModeでWriteフラグを立てる
    If Not clsADORef.ConnectMode And adModeWrite Then
        'Writeフラグが立っていなかったので立てる
        clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    End If
    'SQL実行、本番ではトランザクション有効にしたほうがいいかも
    isCollect = clsADORef.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        DebugMsgWithTime "IncludeZaikoSH_to_TempTable: Do sql fail..."
        MsgBox "IncludeZaikoSH_to_TempTable でSQLの実行に失敗しました。" & err.Description
        GoTo CloseAndExit
    End If
    '結果の件数を返して終わり
    IncludeZaikoSH_to_TempTable = clsADORef.Affected
    Exit Function
ErrorCatch:
    DebugMsgWithTime "IncludeZaikoSH_to_TempTable code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Set dicReplaceParm = Nothing
    Exit Function
End Function
'''Select Distinct した結果を一時テーブルに流し込むメソッド
'''DBPathはメソッド内で変更した後、デフォルトに戻してしまうので、後でADOHanle使用する際には改めて設定する事
'''ファイル名、テーブル名はConstで決め打ちしている
'''args
'''strargDistinctFieldName              Distinctの対象にしたいフィールド名
Private Function SELECT_ToTempTable(strargDistinctFieldName As String) As Boolean
    On Error GoTo ErrorCatch
    'インスタンス変数の初期化状況のチェック
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "SELECT_ToTempTable : inctance not Initialized"
        GoTo CloseAndExit
        SELECT_ToTempTable = False
        Exit Function
    End If
    'まずはDBPath等をデフォルトに設定してやる
    Call clsADORef.SetDBPathandFilenameDefault
    'DBfilenameのみ一時テーブル置き場のものに変える
    clsADORef.DBFileName = TEMP_DB_FILENAME
    '最初に既存の一時テーブルを削除する
    Call clsADORef.DropTable(INV_CONST.T_INV_SELECT_TEMP)
    'SQL組立
    'パラメータ埋め込み用Dictionary
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    dicReplaceParm.Add 0, strargDistinctFieldName
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(INV_CONST.SQL_INV_SELECT_DISTINCT_TO_TEMPTABLE_0FieldName, dicReplaceParm)
    'SQL実行前にConnectModeでWriteフラグが立っているか調べ、たっていなかったら立てる
    If Not clsADORef.ConnectMode And adModeWrite Then
        'Writeフラグが立っていなかったので立てる
        clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    End If
    'SQL実行
    Dim isCollect As Boolean
    isCollect = clsADORef.Do_SQL_with_NO_Transaction
    If Not isCollect Then
        DebugMsgWithTime "SELECT_ToTempTable: Do sql fail..."
        SELECT_ToTempTable = False
        Exit Function
    End If
    DebugMsgWithTime "SELECT_To_TempTable Affected: " & clsADORef.Affected
    'DBPath等を元に戻してやる
    clsADORef.SetDBPathandFilenameDefault
    SELECT_ToTempTable = True
    GoTo CloseAndExit
    Exit Function
ErrorCatch:
    DebugMsgWithTime "SELECT_ToTempTable code: " & err.Number & " Description: " & err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set dicReplaceParm = Nothing
    Exit Function
End Function
'''デイリー棚卸のCSVファイルから重複なしの手配コード先頭n文字のリストを受け取る関数
'''戻り値 string()      String1次元配列で手配コードの先頭n文字（重複なし）が入っている
'''args
'''strargDailyTanaCSVFullPath           デイリー棚卸のCSVファイルのフルパスを渡す。IEの解析が済むまではダイアログを表示し、選択する
'''longargDigitNum                      LeftN文字の文字数を設定、1以上
Private Function ReturnNdigitTehaiCodeFromCSV(strargDailyTanaCSVFullPath As String, longargDigitNum As Long) As String()
    On Error GoTo ErrorCatch
    'インスタンス変数初期化状況チェック
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        DebugMsgWithTime "ReturnNdigitTehaiCodeFromCSV : instance not Initialized"
        ReturnNdigitTehaiCodeFromCSV = Array("")
        GoTo CloseAndExit
        Exit Function
    End If
    '桁数チェック
    If longargDigitNum < 1 Then
        '桁数が1未満
        DebugMsgWithTime "ReturnNdigitTehaiCodeFromCSV: digit num under 1"
        GoTo CloseAndExit
        Exit Function
    End If
    'CSVファイル名のチェック、しばらくはダイアログで選択してもらうことにする
    If strargDailyTanaCSVFullPath = "" Then
        '引数でCSVのファイル名が指定されなかった場合
        'カレントディレクトリをダウンロードディレクトリに移動する
        Call ChCurrentDirW(GetDownloadPath)
        strargDailyTanaCSVFullPath = CStr(Application.GetOpenFilename("CSVファイル,*.csv", 1, "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"))
    End If
    'CSVファイルにフィルタ処理かけた後のファイル名とテーブル名を受け取る
    Dim arrstrFilterdFileandTableName() As String
    arrstrFilterdFileandTableName = GetFilterRangeNameFromExcel(strargDailyTanaCSVFullPath)
    '置換用Dic作成
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    dicReplaceParm.Add 0, arrstrFilterdFileandTableName(0, 1)
    dicReplaceParm.Add 1, CStr(longargDigitNum)
    'SQL文作成
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(INV_CONST.SQL_INV_TEHAICODE_n_0TableName_1DigitNum, dicReplaceParm)
    'AdoのクラスのプロパティにCSVのパスとファイル名を設定
    Dim fso4Digit As FileSystemObject
    Set fso4Digit = New FileSystemObject
    clsADORef.DBPath = fso4Digit.GetParentFolderName(arrstrFilterdFileandTableName(0, 0))
    clsADORef.DBFileName = fso4Digit.GetFileName(arrstrFilterdFileandTableName(0, 0))
    'SQL実行
    Call clsADORef.Do_SQL_with_NO_Transaction
    If clsADORef.RecordCount <= 0 Then
        '結果が0件だったにょ
        DebugMsgWithTime "ReturnNdigitTehaiCodeFromCSV: no entry found"
        GoTo CloseAndExit
        Exit Function
    End If
    Dim arrstr4Digit() As String
    ReDim arrstr4Digit(clsADORef.RecordCount - 1)
    clsADORef.RS.MoveFirst
    Dim longArrayRowCounter As Long
    longArrayRowCounter = 0
    Do Until clsADORef.RS.EOF = True
        arrstr4Digit(longArrayRowCounter) = CStr(clsADORef.RS.Fields(0).Value)
        clsADORef.RS.MoveNext
        longArrayRowCounter = longArrayRowCounter + 1
    Loop
    'ADOHandleの接続を切断しないとファイル削除が失敗してしまう
    Call clsADORef.CloseClassConnection
    '変換後のファイルを削除
    Kill arrstrFilterdFileandTableName(0, 0)
    ReturnNdigitTehaiCodeFromCSV = arrstr4Digit
    GoTo CloseAndExit
    Exit Function
ErrorCatch:
    DebugMsgWithTime "Return4digitTtehaiCodeFromCSV code: " & err.Number & " Description: " & err.Description
CloseAndExit:
    Set fso4Digit = Nothing
    Set dicReplaceParm = Nothing
    Exit Function
End Function
'''T_INV_CSVの内容を（当初は現品残のみ）棚卸CSVに反映させる
'''Return String    オリジナルファイルをリネームしたバックアップファイルのフルパス
'''                 目的のファイルは、引数で与えられたファイルと同名とする
'''args
'''strargSHCSVFullPath  処理対象のCSV(xls)ファイル
'''strargEndDay         棚卸締切日、Right10のやつ
Public Function SetDBDatatoTanaCSV(strargSHCSVFullPath As String, strargEndDay As String) As String
    On Error GoTo ErrorCatch
    '共有インスタンスの初期化状況を確認
    Dim isInitialized As Boolean
    isInitialized = CheckInstance
    If Not isInitialized Then
        '初期化チェック失敗
        DebugMsgWithTime "SetDBDatatoTanaCSV : share instance Initialize fail..."
        SetDBDatatoTanaCSV = ""
        GoTo CloseAndExit
    End If
    'ファイルの存在確認
    Dim fsoSetCSV As FileSystemObject
    Set fsoSetCSV = New FileSystemObject
    If Not fsoSetCSV.FileExists(strargSHCSVFullPath) Then
        MsgBox "指定のファイルが見つかりませんでした: " & strargSHCSVFullPath
        DebugMsgWithTime "SetDBDatatoTanaCSV : file not found: " & strargSHCSVFullPath
        SetDBDatatoTanaCSV = ""
        GoTo CloseAndExit
    End If
    'EndDay引数の空チェック
    If strargEndDay = "" Then
        MsgBox "棚卸締切日が空でした"
        DebugMsgWithTime "SetDBDatatoTanaCSV : EndDay is empty"
        SetDBDatatoTanaCSV = ""
        GoTo CloseAndExit
    End If
    '引数のシートのEndDayを受け取り、引数のEndDayと比較する
    Dim strSHEndDay As String
    strSHEndDay = GetEndDayFromTanaCSV(strargSHCSVFullPath)
    If strSHEndDay <> strargEndDay Then
        DebugMsgWithTime "SetDBDatatoTanaCSV : End Day dif arg and SHFile: " & strargSHCSVFullPath
        '引数のEndDayとシートのEndDayが違った
        Dim longReturnVal As Long
        longReturnVal = MsgBox("保存先のシートとBINカードチェック画面の棚卸締切日が異なります。このまま続行しますか？", vbYesNo)
        If longReturnVal = vbNo Then
            'いいえを押されたっぽい
            MsgBox "処理を中断します"
            SetDBDatatoTanaCSV = ""
            GoTo CloseAndExit
        End If
    End If
    'CSVファイルにフィルタ処理をして、フルパス名とテーブル名を得る
    Dim strarrSetCSV() As String
    strarrSetCSV = GetFilterRangeNameFromExcel(strargSHCSVFullPath)
    If Not IsRedim(strarrSetCSV) Then
        '初期化されてない配列が返ってきた、失敗してるっぽい
        DebugMsgWithTime "SetDBDatatoTanaCSV : fail convert CSV to xls"
        MsgBox "CSVファイル読み込み中にエラーが発生しました"
        SetDBDatatoTanaCSV = ""
        GoTo CloseAndExit
    End If
    'SQL組み立て開始
''{0}    (Sheet Table Name)
''{1}    SHCSV
''{2}    T_INV_CSV
''{3}    (After In Word Default DB)
''{4}    棚卸締切日
''{5}    (2022/01/12 EndDay)
''{6}    TCSVTana
''{7}    手配コード
''{8}    現品残
'Public Const SQL_INV_DB_TO_CSV As String = "UPDATE [{0}] AS {1}" & vbCrLf &
    Dim dicReplaceSetCSV As Dictionary
    Set dicReplaceSetCSV = New Dictionary
    dicReplaceSetCSV.Add 0, strarrSetCSV(0, 1)
    dicReplaceSetCSV.Add 1, clsEnumRef.SQL_INV_Alias(SHCSV_Alias_sia)
    dicReplaceSetCSV.Add 2, INV_CONST.T_INV_CSV
    '一旦clsADOをデフォルトディレクトリへ
    clsADORef.SetDBPathandFilenameDefault
    dicReplaceSetCSV.Add 3, clsSQLBcRef.CreateAfterIN_WordFromSHFullPath(fsoSetCSV.BuildPath(clsADORef.DBPath, clsADORef.DBFileName), clsEnumRef)
    dicReplaceSetCSV.Add 4, clsEnumRef.CSVTanafield(F_EndDay_ICS)
    dicReplaceSetCSV.Add 5, strargEndDay
    dicReplaceSetCSV.Add 6, clsEnumRef.SQL_INV_Alias(TanaCSV_Alias_sia)
    dicReplaceSetCSV.Add 7, clsEnumRef.CSVTanafield(F_Tehai_Code_ICS)
    dicReplaceSetCSV.Add 8, clsEnumRef.CSVTanafield(F_Available_ICS)
    '置換実行、SQL設定
    clsADORef.SQL = clsSQLBcRef.ReplaceParm(INV_CONST.SQL_INV_DB_TO_CSV, dicReplaceSetCSV)
    'DBPath,DBFilenameをフィルタ処理後のxlsのものに
    clsADORef.DBPath = fsoSetCSV.GetParentFolderName(strarrSetCSV(0, 0))
    clsADORef.DBFileName = fsoSetCSV.GetFileName(strarrSetCSV(0, 0))
    'Writeフラグを立てる
    clsADORef.ConnectMode = clsADORef.ConnectMode Or adModeWrite
    'SQL実行
    Dim isCollect As Boolean
    isCollect = clsADORef.Do_SQL_with_NO_Transaction
    'Writeフラグ下げる
    clsADORef.ConnectMode = clsADORef.ConnectMode And Not adModeWrite
    If Not isCollect Then
        DebugMsgWithTime "SetDBDatatoTanaCSV : do sql fail..."
        MsgBox "SetDBDatatoTanaCSV でSQL実行失敗したようです"
        SetDBDatatoTanaCSV = ""
        GoTo CloseAndExit
    End If
    'ファイルリネーム関係の処理
    'オリジナルファイルの名前を変更(バックアップ)
    Dim strarrNewFileName(3) As String
    strarrNewFileName(0) = fsoSetCSV.GetBaseName(strargSHCSVFullPath)
    strarrNewFileName(1) = GetTimeForFileNameWithMilliSec
    strarrNewFileName(2) = "."
    strarrNewFileName(3) = fsoSetCSV.GetExtensionName(strargSHCSVFullPath)
    '保存前のファイルのFileオブジェクトを取得
    Dim fileOrigin As File
    Set fileOrigin = fsoSetCSV.GetFile(strargSHCSVFullPath)
    'リネーム実行
    fileOrigin.Name = Join(strarrNewFileName, "")
    '関数の戻り値としてバックアップファイルの名前をセット
    SetDBDatatoTanaCSV = fileOrigin.Name
    '次に処理後のファイルを引数のフルパスで保存するが、拡張子によって処理を分岐する必要あり
    isCollect = SavesomefileToExelFile(strarrSetCSV(0, 0), strargSHCSVFullPath)
    If Not isCollect Then
        'ファイル変換中にエラー発生
        DebugMsgWithTime "SetDBDatatoTanaCSV : error in exchanging"
        SetDBDatatoTanaCSV = ""
        GoTo CloseAndExit
    End If
    '変換終了したら途中の中間フィルタ処理ファイルは削除する
    Kill strarrSetCSV(0, 0)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SetDBDatatoTanaCSV code: " & err.Number & " Description: " & err.Description
    SetDBDatatoTanaCSV = ""
    GoTo CloseAndExit
CloseAndExit:
    Set fsoSetCSV = Nothing
    Exit Function
End Function
'''与えられたExcel対応ファイルを、指定されたExcel対応ファイルに別名保存する
'''Return   Bool    成功したらTrue、それ以外はFalse
'''args
'''strargOriginFile             変換(リネーム)元のファイル
'''strargSavetoNewFileFullPath  変換(リネーム)先のファイル
'''                             2つのファイル名は別の名前にすること
Private Function SavesomefileToExelFile(strargOriginFile As String, strargSavetoNewFileFullPath As String) As Boolean
    On Error GoTo ErrorCatch
    'ファイル存在確認
    Dim fsoSaveTo As FileSystemObject
    Set fsoSaveTo = New FileSystemObject
    If strargOriginFile = strargSavetoNewFileFullPath Then
        DebugMsgWithTime "SavesomefileToExelFile before and after file path is same"
        MsgBox "変換前と変換後のファイルパスが同一です"
        SavesomefileToExelFile = False
        GoTo CloseAndExit
        Exit Function
    End If
    If Not fsoSaveTo.FileExists(strargOriginFile) Then
        DebugMsgWithTime "SavesomefileToExelFile : origin file not found: " & strargOriginFile
        MsgBox "オリジナルファイルが見つかりません。" & strargOriginFile
        SavesomefileToExelFile = False
        GoTo CloseAndExit
        Exit Function
    End If
    If fsoSaveTo.FileExists(strargSavetoNewFileFullPath) Then
        '変換後のファイルが存在した
        Dim longMsgBoxReturn As Long
        '削除の確認
        longMsgBoxReturn = MsgBox("変換後のファイルが既に存在します。消去してもいいですか？", vbYesNo)
        If longMsgBoxReturn = vbYes Then
            'Yes
            'ファイル消去実行
            Kill strargSavetoNewFileFullPath
        Else
            'NO
            MsgBox "処理を中断します"
            SavesomefileToExelFile = False
            GoTo CloseAndExit
            Exit Function
        End If
    End If
    '変換前のWorkBookオブジェクトの取得
    Dim wkbSaveTo As Workbook
    '読み取り専用で開く
    Set wkbSaveTo = objExcelRef.Workbooks.Open(strargOriginFile, , True)
    '変換後の拡張子により処理を分岐
    Select Case LCase(fsoSaveTo.GetExtensionName(strargSavetoNewFileFullPath))
    Case clsEnumRef.DBFileExetension(csv_dext)
        'CSV
        'このファイル形式だけシートごとの保存になるので、(1)のシートで決め打ちする
        Dim shSeveTo As Worksheet
        Set shSeveTo = wkbSaveTo.Sheets(1)
        shSeveTo.SaveAs strargSavetoNewFileFullPath, xlCSV
        SavesomefileToExelFile = True
        wkbSaveTo.Saved = True
        GoTo CloseAndExit
    Case clsEnumRef.DBFileExetension(xls_dext)
        'xls
        wkbSaveTo.SaveAs strargSavetoNewFileFullPath, XlFileFormat.xlWorkbookNormal
        SavesomefileToExelFile = True
        GoTo CloseAndExit
    Case clsEnumRef.DBFileExetension(xlsb_dext)
        'xlsb
        wkbSaveTo.SaveAs strargSavetoNewFileFullPath, XlFileFormat.xlExcel12
        SavesomefileToExelFile = True
        GoTo CloseAndExit
    Case clsEnumRef.DBFileExetension(xlsm_dext)
        'xlsm
        wkbSaveTo.SaveAs strargSavetoNewFileFullPath, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
        SavesomefileToExelFile = True
        GoTo CloseAndExit
    Case clsEnumRef.DBFileExetension(xlsx_dext)
        'xlsx
        wkbSaveTo.SaveAs strargSavetoNewFileFullPath, XlFileFormat.xlWorkbookDefault
        SavesomefileToExelFile = True
        GoTo CloseAndExit
    End Select
    GoTo CloseAndExit
    Exit Function
ErrorCatch:
    DebugMsgWithTime "SavesomefileToExelFile code: " & err.Number & " Description: " & err.Description
    SavesomefileToExelFile = False
    GoTo CloseAndExit
CloseAndExit:
    Set shSeveTo = Nothing
    If Not wkbSaveTo Is Nothing Then
        wkbSaveTo.Saved = True
        wkbSaveTo.Close
        Set wkbSaveTo = Nothing
    End If
    Exit Function
End Function
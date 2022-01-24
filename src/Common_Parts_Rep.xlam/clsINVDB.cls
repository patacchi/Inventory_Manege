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
            GoTo CloseAndExit
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
        Err.Clear
        Set shtFilter = wkbFilter.Worksheets(strargSheetName)
        'Err.Numberが0以外の時は処理を中断
        If Err.Number <> 0 Then
            DebugMsgWithTime "GetFilterRangeNameFromExcel: sheet not found: " & strargSheetName
            GoTo CloseAndExit
            Exit Function
        End If
        On Error GoTo ErrorCatch
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
           GoTo CloseAndExit
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
        GoTo CloseAndExit
        Exit Function
    Case Else
        'Excelファイルじゃなかった場合
        DebugMsgWithTime "SetExcelAutoFormat: not excel file"
        GoTo CloseAndExit
    End Select
ErrorCatch:
    DebugMsgWithTime "GetFilterRangeNameFromExcel code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
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
    Dim EnumValue As clsEnum
    Set EnumValue = CreateclsEnum
    If Not LCase(fsoCSV.GetExtensionName(strargCSVFileFullPath)) = EnumValue.DBFileExetension(csv_dext) Then
        '拡張子がcsvじゃなかった
        DebugMsgWithTime "CSVFileSaveAsxlsx: not csv files (extention)"
        GoTo CloseAndExit
        Exit Function
    End If
    'CSVファイルを開き、xlsx形式で保存する
    If objExcel Is Nothing Then
        'クラスのobjExcelが初期化されていなかったらここで初期化する
        Set objExcel = New Excel.Application
    End If
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
    objExcel.CutCopyMode = False
    '改めてxlsxファイルを保存
    wkbXlsx.Save
    wkbXlsx.Close
    Set wkbXlsx = Nothing
    CSVFileSaveAsXlsxFileandPasteValues = strxlsxFileName
    GoTo CloseAndExit
ErrorCatch:
    CSVFileSaveAsXlsxFileandPasteValues = ""
    DebugMsgWithTime "CSVFileSaveAsXlsxFileandPasteValues code: " & Err.Number & " Description: " & Err.Description
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
    Set EnumValue = Nothing
    Set fsoCSV = Nothing
    Exit Function
End Function
'''Author Daisuke Oota 2022_01_20
'''在庫情報シートのフルパスを引数として、INVパーツマスターDBに登録するプロシージャ
'''args
'''strargZaikoSHFullPath    在庫情報シートのフルパス
'''clsADOUpsert             ADOHandleのインスタンス、繰り返し処理をすることになると思うので、このメソッド以前で宣言し、共有する
'''                         DBPath,DBFilenameは引数で渡されるクラスのインスタンスのプロパティで設定することとする
'''clsINVDBUpsert           INVDBクラスのインスタンスを指定する。このメソッド以前に作成していたら再利用する（コンストラクタの動作が重いため）
'''clsEnumUpsert            Enunクラスのインスタンスを指定する、できれば事前に一度だけ作成したのを使いまわしたい
'''戻り値 long      変更があったレコード数を返す。0以下だと何かしらがあった？
Public Function UpsertINVPartsMasterfromZaikoSH(strargZaikoSHFullPath As String, _
ByRef clsINVDBUpsert As clsINVDB, ByRef clsADOUpsert As clsADOHandle, ByRef clsEnumUpsert As clsEnum) As Long
    On Error GoTo ErrorCatch
    'クラスの存在確認
    'clsINVDBUpsert
    If clsINVDBUpsert Is Nothing Then
        '初期化されてないクラスインスタンスが渡された
        DebugMsgWithTime "UpsertINVPartsMasterfromZaikoSH: Class instance is nothing"
        UpsertINVPartsMasterfromZaikoSH = 0
        GoTo CloseAndExit
        Exit Function
    End If
    'clsADOHandle
    If clsADOUpsert Is Nothing Then
        DebugMsgWithTime "UpsertINVPartsMasterfromZaikoSH: warning! ADOUpsert instance is nothing.will delay..."
        Set clsADOUpsert = CreateclsADOHandleInstance
    End If
    'clsEnum
    If clsEnumUpsert Is Nothing Then
        DebugMsgWithTime "UpsertINVPartsMasterFromZaikoSH: warning! Enum instance is nothing.will delay..."
        Set clsEnumUpsert = CreateclsEnum
    End If
'#If DebugDB = 1 Then
'    'DebugDBモードの時
'    strargZaikoSHFullPath = Application.GetOpenFilename()
'#End If
    'ファイルの存在確認
    Dim fsoUpsertPartsMaster As FileSystemObject
    Set fsoUpsertPartsMaster = New FileSystemObject
    If Not fsoUpsertPartsMaster.FileExists(strargZaikoSHFullPath) Then
        '引数で指定したファイルが存在しなかった場合
        DebugMsgWithTime "UpsertINVPartsMaster file not found: " & strargZaikoSHFullPath
        UpsertINVPartsMasterfromZaikoSH = 0
        GoTo CloseAndExit
    End If
    Dim strFilterdZaikoSh() As String
    'フィルタ処理後のファイル名を受け取る。この時、シート名を明示的にINT_CONSTから引っ張り指定してやる
    strFilterdZaikoSh = clsINVDBUpsert.GetFilterRangeNameFromExcel(strargZaikoSHFullPath, INV_CONST.INV_SH_ZAIKO_NAME)
    '処理後に初期化されていない配列が帰ってきたら、在庫情報シートが存在しない可能性が高いので処理を中断する
    If Not IsRedim(strFilterdZaikoSh) Then
        DebugMsgWithTime "UpserINVPartsMaster: not found zako_info sheets origin file: " & strargZaikoSHFullPath
        UpsertINVPartsMasterfromZaikoSH = 0
        GoTo CloseAndExit
    End If
    '''ここまでで在庫情報シートが確定したので、在庫情報からDBに登録するメソッドで処理を行う
    Dim longResultAffected As Long
    'まずは在庫情報シートをTempTableに格納する（直接Excelファイル参照だとうまくいかなかった）
    longResultAffected = IncludeZaikoSH_to_TempTable(strFilterdZaikoSh(0, 0), strFilterdZaikoSh(0, 1), clsADOUpsert, clsEnumUpsert)
    UpsertINVPartsMasterfromZaikoSH = longResultAffected
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "UpsertINBPartMasterFromZaikoSH code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set fsoUpsertPartsMaster = Nothing
    'INVDBインスタンスはByRefで他の物を再利用しているので解放禁止
'    If Not INVDBupsertINVPartMaster Is Nothing Then
'        'INVDB変数がNothingだった場合
'        Set INVDBupsertINVPartMaster = Nothing
'    End If
    Exit Function
End Function
'''Author Daisuke oota 2022_01_21
'''在庫情報シートのフルパスを引数として、DBの一時テーブルに登録するメソッド
'''args
'''strargFilterdZaikoSHFullPath         フィルタ処理済みのブック名のフルパス
'''strargTableName                      在庫情報シートのSelect 文 Fromに入るテーブル名
'''clsADOUpsert                         ADOHandleのインスタンス、繰り返し処理をすることになると思うので、このメソッド以前で宣言し、共有する
'''                                     DBPath,DBFilenameは引数で渡されるクラスのインスタンスのプロパティで設定することとする
'''戻り値 long      影響のあった行数を返す
Private Function IncludeZaikoSH_to_TempTable(strargFilterdZaikoSHFullPath As String, strargTableName As String, _
ByRef clsADOUpsert As clsADOHandle, ByRef EnumUpsert As clsEnum) As Long
    On Error GoTo ErrorCatch
    'AdoHandleのクラス初期化状況を確認
    If clsADOUpsert Is Nothing Then
        DebugMsgWithTime "IncludeZaikoSH_to_TempTable: warning! clsADOHandle instance Nothing. will delay cause create new class"
        Set clsADOUpsert = CreateclsADOHandleInstance
        IncludeZaikoSH_to_TempTable = 0
        GoTo CloseAndExit
        Exit Function
    End If
    'このメソッドでは一時テーブルDBを扱うので、明示的にファイルを指定してやる
    clsADOUpsert.DBFileName = PublicConst.TEMP_DB_FILENAME
    'ファイル、テーブルの空データチェック
    If strargFilterdZaikoSHFullPath = "" Or strargTableName = "" Then
        DebugMsgWithTime "IncludeZaikoSH_to_TempTable: Zaiko sheet filename or table name is empty"
        IncludeZaikoSH_to_TempTable = 0
        GoTo CloseAndExit
        Exit Function
    End If
    '在庫情報シートの内容を一時テーブルに格納する
    'まずは最初に一時テーブルを消去してやる
    Call clsADOUpsert.DropTable(INV_CONST.T_INV_TEMP, PublicConst.TEMP_DB_FILENAME)
    'SQL文の組み立て
    'INV_CONST.SQL_INV_SH_TO_DB_TEMPTABLE_0Table_1INword
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = CreateclsSQLStringBuilder
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    dicReplaceParm.Add 0, strargTableName
    dicReplaceParm.Add 1, CreateAfterIN_WordFromSHFullPath(strargFilterdZaikoSHFullPath, EnumUpsert)
    clsADOUpsert.SQL = sqlBC.ReplaceParm(INV_CONST.SQL_INV_SH_TO_DB_TEMPTABLE_0Table_1INword, dicReplaceParm)
    'SQL実行、本番ではトランザクション有効にしたほうがいいかも
    Call clsADOUpsert.Do_SQL_with_NO_Transaction
    '結果の件数を返して終わり
    IncludeZaikoSH_to_TempTable = clsADOUpsert.Affected
    Exit Function
ErrorCatch:
    DebugMsgWithTime "IncludeZaikoSH_to_TempTable code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Set sqlBC = Nothing
    Set dicReplaceParm = Nothing
    Exit Function
End Function
'''ExcelファイルフルパスからIN句の後に続く文字列を作成する
'''戻り値 String        Select FROM の IN句 ""の後に付く文字列ex)[Excel 12.0;DATABASE=R:\Tmp\Patacchi\Test Dir\Zaiko_0_Local.xls;]
'''args
'''EnumINWord          Enumの共有インスタンス
Public Function CreateAfterIN_WordFromSHFullPath(strargSHFullPath As String, ByRef EnumInWord As clsEnum) As String
    If strargSHFullPath = "" Then
        DebugMsgWithTime "CreateAfterINword: sheet file name empty"
    End If
    Dim dicExttoEnumValu As Dictionary
    '置換用Dic作成
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    '{0}がファイルフルパスなのは確定してるのでここで入れてしまう
    dicReplaceParm.Add 0, strargSHFullPath
    'ファイルの拡張子により処理を分岐
    Dim fsoINWord As FileSystemObject
    Set fsoINWord = New FileSystemObject
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = CreateclsSQLStringBuilder
    Select Case LCase(fsoINWord.GetExtensionName(strargSHFullPath))
    Case EnumInWord.DBFileExetension(accdb_dext)
        'accdb
        CreateAfterIN_WordFromSHFullPath = sqlBC.ReplaceParm(INV_CONST.SQL_AFTER_IN_ACCDB_0FullPath, dicReplaceParm)
    Case EnumInWord.DBFileExetension(xlam_dext), EnumInWord.DBFileExetension(xlsm_dext)
        'xlsm (xlam・・は使わないと思う)
        CreateAfterIN_WordFromSHFullPath = sqlBC.ReplaceParm(INV_CONST.SQL_AFTER_IN_XLSM_0FullPath, dicReplaceParm)
    Case EnumInWord.DBFileExetension(xlsx_dext)
        'xlsx
        CreateAfterIN_WordFromSHFullPath = sqlBC.ReplaceParm(INV_CONST.SQL_AFTER_IN_XLSX_0FullPath, dicReplaceParm)
    Case EnumInWord.DBFileExetension(xlsb_dext)
        'xlsb
        CreateAfterIN_WordFromSHFullPath = sqlBC.ReplaceParm(INV_CONST.SQL_AFTER_IN_XLSB_0FullPath, dicReplaceParm)
    Case EnumInWord.DBFileExetension(xls_dext)
        'xls
        CreateAfterIN_WordFromSHFullPath = sqlBC.ReplaceParm(INV_CONST.SQL_AFTER_IN_XLS_0FullPath, dicReplaceParm)
    Case Else
        'それ以外は（CSVも含めて）対応なしの予定
        DebugMsgWithTime "CreateAfterINWordFromSHFullPath: no support extention: " & LCase(fsoINWord.GetExtensionName(strargSHFullPath))
        CreateAfterIN_WordFromSHFullPath = ""
    End Select
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "CreateAfterINWordFromSHFullPath code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set dicReplaceParm = Nothing
    Set fsoINWord = Nothing
    Exit Function
End Function
'''デイリー棚卸のCSVファイルから重複なしの手配コード先頭4文字のリストを受け取る関数
'''戻り値 string()      String1次元配列で手配コードの先頭4文字（重複なし）が入っている
'''args
'''strargDailyTanaCSVFullPath           デイリー棚卸のCSVファイルのフルパスを渡す。IEの解析が済むまではダイアログを表示し、選択する
'''clsAdo4Digit                         この3個はインスタンスをプロシージャ間で共有し、コンストラクタの動作を少しでも軽量化したい
'''clsINVDBdigit
'''clsEnumDigit
Public Function Return4digitTehaiCodeFromCSV(strargDailyTanaCSVFullPath As String, _
ByRef clsAdo4Digit As clsADOHandle, ByRef clsINVDBdigit As clsINVDB, ByRef clsEnumDigit As clsEnum) As String()
    On Error GoTo ErrorCatch
    'CSVファイル名のチェック、しばらくはダイアログで選択してもらうことにする
    If strargDailyTanaCSVFullPath = "" Then
        '引数でCSVのファイル名が指定されなかった場合
        'カレントディレクトリをダウンロードディレクトリに移動する
#If DebugDB Then
        'ローカル試験用
'        'カレントディレクトリをDBディレクトリにする
'        Call ChCurrentDirW(clsAdo4Digit.DBPath)
        'カレントディレクトリをダウンロードディレクトリに移動する
        Call ChCurrentDirW(GetDownloadPath)
#Else
        '本番環境用
        'カレントディレクトリをダウンロードディレクトリに移動する
        Call ChCurrentDirW(GetDownloadPath)
#End If
        strargDailyTanaCSVFullPath = CStr(Application.GetOpenFilename("CSVファイル,*.csv", 1, "デイリー棚卸でダウンロードしたCSVファイルを選択して下さい"))
    End If
    'CSVファイルにフィルタ処理かけた後のファイル名とテーブル名を受け取る
    Dim arrstrFilterdFileandTableName() As String
    arrstrFilterdFileandTableName = Me.GetFilterRangeNameFromExcel(strargDailyTanaCSVFullPath)
    '置換用Dic作成
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    dicReplaceParm.Add 0, arrstrFilterdFileandTableName(0, 1)
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = CreateclsSQLStringBuilder
    'SQL文作成
    clsAdo4Digit.SQL = sqlBC.ReplaceParm(INV_CONST.SQL_INV_TEHAICODE_FIRST4_0TableName, dicReplaceParm)
    'AdoのクラスのプロパティにCSVのパスとファイル名を設定
    Dim fso4Digit As FileSystemObject
    Set fso4Digit = New FileSystemObject
    clsAdo4Digit.DBPath = fso4Digit.GetParentFolderName(arrstrFilterdFileandTableName(0, 0))
    clsAdo4Digit.DBFileName = fso4Digit.GetFileName(arrstrFilterdFileandTableName(0, 0))
    'SQL実行
    Call clsAdo4Digit.Do_SQL_with_NO_Transaction
    If clsAdo4Digit.RecordCount <= 0 Then
        '結果が0件だったにょ
        DebugMsgWithTime "Return4digitTehaiCodeFromCSV: no entry found"
        GoTo CloseAndExit
        Exit Function
    End If
    Dim arrstr4Digit() As String
    ReDim arrstr4Digit(clsAdo4Digit.RecordCount - 1)
    clsAdo4Digit.RS.MoveFirst
    Dim longArrayRowCounter As Long
    longArrayRowCounter = 0
    Do Until clsAdo4Digit.RS.EOF = True
        arrstr4Digit(longArrayRowCounter) = CStr(clsAdo4Digit.RS.Fields(0).Value)
        clsAdo4Digit.RS.MoveNext
        longArrayRowCounter = longArrayRowCounter + 1
    Loop
    Call clsAdo4Digit.CloseClassConnection
    '変換後のファイルを削除
    Kill arrstrFilterdFileandTableName(0, 0)
    Return4digitTehaiCodeFromCSV = arrstr4Digit
    GoTo CloseAndExit
    Exit Function
ErrorCatch:
    DebugMsgWithTime "Return4digitTtehaiCodeFromCSV code: " & Err.Number & " Description: " & Err.Description
CloseAndExit:
    Set fso4Digit = Nothing
    Set dicReplaceParm = Nothing
    Set sqlBC = Nothing
    Exit Function
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsADOHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'''ADODBによりDBを操作するクラス
'2021_10_10 IE連携のために新規作成 Patacchi
'DB接続関係の定数設定
Const constDatabasePath             As String = "C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\bin\Inventory_DB"      'データベースディレクトリ
#If DebugDB Then
'Local環境
Const constDBFILENAME               As String = "INV_Manege_Local.accdb"    '在庫管理ローカル用
#Else
'本番環境
Const constDBFILENAME               As String = "INV_Manege.accdb"          '在庫管理DB
#End If
Const DEFAULT_CONNECTION_MODE       As Long = ConnectModeEnum.adModeRead Or ConnectModeEnum.adModeShareDenyNone
Const constDBPROVIDER_Pri           As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source="
Const CONNECTION_STRING_ACCDB_0DBFullPath           As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0}"       '接続文字列_accdb
Const CONNECTION_STRING_XLSM_0DBFullPath            As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Macro;HDR=Yes;"";"
'xlamはxlsmと同じ接続文字列・・だけどxlamに対して接続する事はないかも・・・
Const CONNECTION_STRING_XLAM_0DBFullPath            As String = PublicConst.DB_FILE_EXETENSION_XLSM
Const CONNECTION_STRING_XLSB_0DBFullPath            As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=Yes;"";"
Const CONNECTION_STRING_XLSX_0DBFullPath            As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 xml;HDR=Yes;"";"
Const CONNECTION_STRING_XLS_0DBFullPath             As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=Yes;"";"
'CSVだけDataSourceがファイルの入ってるディレクトリまでなので注意、ファイル名はFROM のテーブル名として渡す
Const CONNECTION_STRING_CSV_0DBPathOnly             As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Text;HDR=Yes;FMT=Delimited"""
'メンバ変数
Private strDBPath As String
Private strDBFileName As String
Private strConnection As String
Private isConnected As Boolean
Private strSQL As String
Private adoConnection As ADODB.Connection
Private adoCommand As ADODB.Command
Private adoParameters As ADODB.Parameters
Private rsMyRecord As ADODB.Recordset
Private varMyArray As Variant
Private longAffected As Long
Private conMode As ConnectModeEnum
'影響を受けた行数を格納する（UpdateとかInsertとか）
'プロパティ
'DBPath
Property Get DBPath() As String
    If strDBPath = "" Then
        'メンバ変数が初期状態なら定数をセットしてやる
'        strDBPath = constDatabasePath
        Me.DBPath = constDatabasePath
    End If
    DBPath = strDBPath
End Property
Property Let DBPath(ByVal strargDBPath As String)
    If strargDBPath = "" Then
        '引数が空白なら抜ける
        Exit Property
    End If
    '既存のパスと違っていた場合既に接続済みの場合はいったん接続を切断する
    If strargDBPath <> strDBPath Then
        '既存のパスと違っていた場合
        '接続状態により処理を分岐
        If Me.Connected Then
            '接続されていた場合
            '一旦接続を切断する
            '再接続はここではやらない
            Call CloseClassConnection
            'パスをメンバ変数に設定してやる
            strDBPath = strargDBPath
            Exit Property
        Else
            '接続されていなかった場合
            strDBPath = strargDBPath
            Exit Property
        End If
    Else
        '既存のパスと同じだったら何もしない
        Exit Property
    End If
End Property
'Connectionオブジェクト(読み取り専用)
Property Get ConnectionObj() As ADODB.Connection
    Set ConnectionObj = adoConnection
    Exit Property
End Property
'DBFileName
Property Get DBFileName() As String
    If strDBFileName = "" Then
        'ファイル名のメンバ変数が初期状態なら定数をセットしてやる
        strDBFileName = constDBFILENAME
    End If
    DBFileName = strDBFileName
End Property
Property Let DBFileName(ByVal strargDBFileName As String)
    If strargDBFileName = "" Then
        '引数が空だったら抜ける
        Exit Property
    End If
    '既存のファイル名と違っていて、かつ接続済みの場合は一旦接続を切断する
    If strargDBFileName <> strDBFileName Then
        '既存のファイル名と違っていた場合
        '接続状態により処理を分岐
        If Me.Connected Then
            '接続済みだった場合
            '一旦接続を切断する
            Call CloseClassConnection
            'メンバ変数にセットしてやる
            strDBFileName = strargDBFileName
            Exit Property
        Else
            '未接続だった場合
            'そのままメンバ変数にセットしてやる
            strDBFileName = strargDBFileName
            Exit Property
        End If
    Else
        '既存のファイル名と同じだったら何もしない
        Exit Property
    End If
End Property
'ConnectionString
Property Get ConnectionString() As String
    '接続状態により処理を分岐
    If Me.Connected Then
        '接続済みなら既存の接続文字列を返してやる
        ConnectionString = strConnection
        Exit Property
    Else
        '未接続で、更にクラスのメンバ変数が空だったらプロパティのDBPathとDBFilenameから接続文字列を生成してやる
        If strConnection = "" Then
            'クラスのstrConnectionが空だった（未指定状態)
            strConnection = CreateConnectionString(Me.DBPath, Me.DBFileName)
        End If
        ConnectionString = strConnection
        Exit Property
    End If
End Property
Property Let ConnectionString(ByVal strargConnection As String)
    '既に接続済みの場合は変更させない
    If Me.Connected Then
        Exit Property
    End If
    If strargConnection <> "" Then
        strConnection = strargConnection
    End If
End Property
'isConnected
'読み取り専用
Property Get Connected() As Boolean
    '接続の有無の確認方法を変更、外部で接続を閉じられた時の対応のため 2021_12_19 Patacchi
    If Not adoConnection Is Nothing Then
        'ConnectionオブジェクトがNothingではなく
        If adoConnection.State And adStateOpen Then
            'かつOpenフラグが立ってたら接続済みとする
            'adStateConnectingは接続の真っ最中なので違う
            isConnected = True
        Else
            isConnected = False
        End If
        Connected = isConnected
        Exit Property
    End If
    'ここまで抜けてきたら未接続とする
    isConnected = False
    Connected = isConnected
End Property
'SQL
Property Get SQL() As String
    SQL = strSQL
End Property
Property Let SQL(ByVal strargSQL As String)
    If Not strargSQL = "" Then
        strSQL = strargSQL
    End If
End Property
'RecordSet
'読み取り専用？じゃなくなったかも・・・
Property Get RS() As ADODB.Recordset
    Set RS = rsMyRecord
End Property
Property Set RS(argRS As ADODB.Recordset)
    Set rsMyRecord = argRS
    Exit Property
End Property
'RA_Array
'読み取り専用
Property Get RS_Array(Optional NoTitle As Boolean = False) As Variant
    If NoTitle Then
        'タイトル無しの場合は2行目からの配列を再作成し、返す
        If Me.RecordCount < 1 Or Not IsRedim(varMyArray) Then
            'レコードカウントが1未満の時か、配列が初期化されていないとき
            'データなしだよ通知をして終わり
            RS_Array = Empty
            Exit Property
        End If
        '2行目からのデータ作成開始
        Dim longNotitleArrayRowCounter As Long
        Dim longNotitleArrayColumnCounter As Long
        '結果格納用配列定義
        Dim varResult As Variant
        ReDim varResult(Me.RecordCount - 1, Me.RS.Fields.Count - 1)
        '行数だけ明示的に2行目から始める
        For longNotitleArrayRowCounter = 1 To UBound(varMyArray, 1)
            'フィールド数分ループ
            For longNotitleArrayColumnCounter = LBound(varMyArray, 2) To UBound(varMyArray, 2)
                'ArrayRowCounterが1から始まってるので注意！
                varResult(longNotitleArrayRowCounter - 1, longNotitleArrayColumnCounter) = varMyArray(longNotitleArrayRowCounter, longNotitleArrayColumnCounter)
            Next longNotitleArrayColumnCounter
        Next longNotitleArrayRowCounter
        '結果を返して終了
        RS_Array = varResult
        Exit Property
    Else
        'デフォルトはこっち、タイトル付き
        RS_Array = varMyArray
        Exit Property
    End If
End Property
'RecordCount
'読み取り専用
Property Get RecordCount() As Long
    If rsMyRecord.State = 1 And rsMyRecord.RecordCount > 0 Then
        'RSが開いていて、かつRecordCountが1以上の場合に数を返す
        RecordCount = CLng(rsMyRecord.RecordCount)
        Exit Property
    Else
        If rsMyRecord.State = 0 Then
            'State=0 RSが閉じている
            RecordCount = -1
            Exit Property
        End If
    End If
End Property
'Affected
'読み取り専用
Property Get Affected() As Long
    If longAffected >= 0 Then
        Affected = longAffected
    End If
End Property
'ConnectMode
Property Get ConnectMode() As ConnectModeEnum
    If Not adoConnection Is Nothing Then
        If Me.Connected Then
            '初期化済み、かつ接続済みなら現在のConnectionModeを返す
            conMode = adoConnection.Mode
            ConnectMode = conMode
            Exit Property
        End If
    End If
    If conMode = adModeUnknown Then
        'ConnectionModeが未設定（UnKnown)の場合は初期値を設定してやる
        conMode = DEFAULT_CONNECTION_MODE
        '結果を返す
        ConnectMode = conMode
        Exit Property
    Else
        'それ以外の場合はメンバ変数の値をそのまま返してやる
        ConnectMode = conMode
        Exit Property
    End If
End Property
Property Let ConnectMode(argConnectMode As ConnectModeEnum)
    If adoConnection Is Nothing Then
        'Connectionの初期化がまだならここで新しいインスタンスを割り当てる
        Set adoConnection = New ADODB.Connection
    End If
    If adoConnection.Mode = argConnectMode Then
        '既にコネクションモードが同じだったら何もする必要なし
        DebugMsgWithTime "ConnectMode Let: Already connection mode set to: " & argConnectMode
        Exit Property
    End If
    If Not adoConnection Is Nothing Then
        '初期化済みで
        If adoConnection.Mode <> argConnectMode And Me.Connected Then
            '引数で指定されたModeが違った場合で更に接続済みの場合
            '一旦接続を切断してやる（多分切り替える）
            Me.CloseClassConnection
        End If
    Else
        '初期化がまだならここで新しいインスタンスを割り当てる
        Set adoConnection = New ADODB.Connection
    End If
    'メンバ変数にセットし、ConnectionObjectにもセットしてやる
    conMode = argConnectMode
    adoConnection.Mode = conMode
    Exit Property
End Property
'''コンストラクタ
'''クラス生成時に最初に行う
Private Sub Class_Initialize()
    Call Initialize
End Sub
'''デストラクタ
'''クラス破棄時最後に行う
Private Sub Class_Terminate()
    Call Finalize
End Sub
Private Sub Initialize()
    isConnected = False
    strSQL = ""
End Sub
Private Sub Finalize()
        'Connectionの後始末
    If Not adoConnection Is Nothing Then
        'Nothingじゃなかったら接続を閉じる
        '接続されていなければ閉じる
        If Me.Connected Then
'            Call adoConnection.Close
            Call CloseClassConnection
        End If
        Set adoConnection = Nothing
    End If
    'Command
    Set adoCommand = Nothing
    'Patrameters
    Set adoParameters = Nothing
End Sub
'''DBPathとDBFilenameをクラスデフォルトの値に設定する
'''arg 無し
Public Sub SetDBPathandFilenameDefault()
    Me.DBPath = constDatabasePath
    Me.DBFileName = constDBFILENAME
End Sub
'''SQLをトランザクションなしで実行する
'''
Public Function Do_SQL_with_NO_Transaction(Optional ByVal strargSQL As String, Optional ByVal strargDBPath As String, Optional ByVal strargDBFileName As String) As Boolean
    Dim strlocalSQL As String
    If strargSQL = "" Then
        'SQLの引数が空白だったらプロパティから引っ張る
        strargSQL = Me.SQL
        If strargSQL = "" Then
            'プロパティから引っ張っても空だったら処理しないで抜ける
            DebugMsgWithTime "Do_SQL_With_NO_Transattion SQL String Empty"
            Do_SQL_with_NO_Transaction = False
            Exit Function
        End If
    End If
    'DBの接続まで行う（定型処理）
    'これが終わればSQLを受付可能な状態になってるはず
    Dim isCollect As Boolean
    isCollect = DBReady(strargDBPath, strargDBFileName)
    If Not isCollect Then
        'DB開いてる最中に何か失敗した
        DebugMsgWithTime "Do_DQL_with_NO_Transaction fail try open DB..."
        Do_SQL_with_NO_Transaction = False
        Exit Function
    End If
    'SQLをセットし、実行する
    On Error GoTo ErrorCatch
    If adoCommand Is Nothing Then
        'Commandが初期化されてなかったらここで行う
        Set adoCommand = New ADODB.Command
    End If
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = strargSQL
    adoCommand.ActiveConnection = adoConnection
    If rsMyRecord Is Nothing Then
        '結果格納用Recordsetの初期化
        Set rsMyRecord = New ADODB.Recordset
    End If
    Set rsMyRecord = adoCommand.Execute(longAffected)
    If rsMyRecord.State = adStateClosed Then
        'Recordsetが閉じていた場合
        'DROP TABLE , SELECT INTO ... 等が該当？
        If Err.Number = 0 Then
            '正常終了？
            Do_SQL_with_NO_Transaction = True
            ReDim varMyArray(1, 0)
            varMyArray(0, 0) = "正常終了"
            varMyArray(1, 0) = "結果を返さないSQLです。" & longAffected & "件のデータに影響がありました。 正常終了しているようです"
        Else
            ReDim varMyArray(1, 0)
            varMyArray(0, 0) = "異常終了"
            varMyArray(1, 0) = "結果を返さないSQLです。異常終了しているようです"
            Do_SQL_with_NO_Transaction = False
        End If
        Exit Function
    End If
    'rsMyRecordのEOFとBOFが違っていたら中身があるので、配列に格納する
    If Not (rsMyRecord.BOF And rsMyRecord.EOF) Then
        '配列書き出し開始
        isCollect = RecordSet_to_2DArray(rsMyRecord)
    End If
    Set adoCommand = Nothing
    Do_SQL_with_NO_Transaction = True
    Exit Function
ErrorCatch:
    MsgBox "Do_SQL_with_NO_Transaction code: " & Err.Number & vbCrLf & "Description: " & Err.Description
    DebugMsgWithTime "Do_SQL_with_NO_Transaction code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
'''Write Daisuke_Oota
'''RecordSetを二次元配列に書き出す
'''結果はVariant型の配列として書き出す
'''1行目（インデックス0）はタイトル列（もしくはエラーメッセージ？）とする
Private Function RecordSet_to_2DArray(ByRef argRS As ADODB.Recordset) As Boolean
    If Me.RecordCount < 0 And Me.Affected < 0 Then
        'RecordCountとAffectedが0未満の時は何かしらのエラーがあった時なのでエラーがありましたメッセージを返す
        ReDim varMyArray(0, 0)
        varMyArray(0, 0) = "エラーが発生したようです"
        RecordSet_to_2DArray = False
        Exit Function
    End If
    If Me.RecordCount < 0 Then
        'RecordCountが取れてないけどAffectedに件数が入ってる→Updateとかじゃないかな
        ReDim varMyArray(0, 0)
        varMyArray(0, 0) = Me.Affected & " 件のレコードを処理しました。"
        RecordSet_to_2DArray = True
        Exit Function
    End If
    'ここからはRecordSetに中身が入っている
    If argRS.Fields.Count >= 1 Then
        'フィールドがある場合はフィールドリストを配列の1行目に追加する
        '結果格納用配列を初期サイズでRedimする
        '結果格納配列のサイズ調整行数はタイトル分があるので、RecoreCountをそのままセット
        ReDim varMyArray(Me.RecordCount, argRS.Fields.Count - 1)
        Dim longFieldCount As Long
        For longFieldCount = 0 To argRS.Fields.Count - 1
            varMyArray(0, longFieldCount) = argRS(longFieldCount).Name
        Next longFieldCount
    Else
        RecordSet_to_2DArray = False
        Exit Function
    End If
    '次に結果データを配列に格納していく
    Dim longRowCount As Long
    longRowCount = 1
    argRS.MoveFirst
    Do
        For longFieldCount = 0 To argRS.Fields.Count - 1
'            argRS(longFieldCount).Type
            'adUnsighnedTinyInt バイト型 のままだと表示出来ない・・？
            If argRS(longFieldCount).Type = adUnsignedTinyInt Then
                varMyArray(longRowCount, longFieldCount) = CInt(argRS(longFieldCount).Value)
            Else
                varMyArray(longRowCount, longFieldCount) = argRS(longFieldCount).Value
            End If
        Next longFieldCount
        '次のレコードへ
        argRS.MoveNext
        If argRS.EOF Then
            'EOFが来たらそこでループを抜ける
            Exit Do
        End If
        '行カウンタインクリメント
        longRowCount = longRowCount + 1
    Loop
'    '結果書き出し配列のダウンサイズ
'    ReDim varMyArray(longRowCount, argRS.Fields.Count - 1)
    RecordSet_to_2DArray = True
    Exit Function
End Function
'''DBのディレクトリ・ファイル存在確認、接続オープンまで行う（SQL発行する直前まで持っていく）
'''
Private Function DBReady(Optional strargDBPath As String, Optional strargDBFileName As String) As Boolean
    If strargDBPath = "" Then
        'DBファイルディレクトリ引数がからだったらプロパティから引っ張る
        strargDBPath = Me.DBPath
    End If
    'カレントディレクトリをDBディレクトリにする
    Dim isCollect As Boolean
    isCollect = ChCurrentToDBDirectory(strargDBPath)
    If Not isCollect Then
        'ディレクトリ移動時に失敗したっぽい
        DebugMsgWithTime "DBReady fail to move DB Directory"
        DBReady = False
        Exit Function
    End If
    If strargDBFileName = "" Then
        'DBファイル名引数が空だったらプロパティから引っ張る
        strargDBFileName = Me.DBFileName
    End If
    '一時テーブル格納DBの存在チェックと作成
    Call CreateandCheckTempDB
    '接続済みで、クラスのConnectionStringと引数から作成したConnectionStringが違う場合は一旦切断する
    '明示的に指定されたパスとファイル名を優先させるため
    If Me.Connected And Me.ConnectionString <> CreateConnectionString(strargDBPath, strargDBFileName) Then
        '接続済みでかつクラスの接続文字列と引数より生成したConnectionstringが違う場合
        '一旦接続を解除してやる
        Call CloseClassConnection
    End If
    '接続文字列を設定する
    Me.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName)
    'DBファイルの存在有無を確認し、無ければ新規作成する
    isCollect = False
    isCollect = IsDBFileExist(strargDBFileName)
    If Not isCollect Then
        MsgBox "DBファイルが見つからないようなので新規作成します"
        InitialDBCreate (strargDBFileName)
    End If
    'DBへ接続する
    isCollect = OpenConnection(Me.ConnectionString)
    If Not isCollect Then
        '接続に失敗したっぽいのでそのまま処理せず抜ける
        DebugMsgWithTime "DBReady: fail to open DB File"
        DBReady = False
        Exit Function
    End If
    DBReady = True
    Exit Function
End Function
'''接続文字列を作成する
'''与えられた拡張子により接続文字を分岐する
'''Return   string  生成された接続文字列
'''args
'''Optional NoExistCheck    Trueでファイルの存在確認を行わない。ファイル新規作成時にAdox.catalogを使用する場合に、存在しないファイル名を指定する可能性があるため
Public Function CreateConnectionString(Optional ByVal strargDBPath As String, Optional ByVal strargDBFileName As String, Optional NoExistCheck As Boolean = False) As String
    If strargDBPath = "" Then
        '引数でDBパスが指定されていなかったらプロパティから引っ張る
        strargDBPath = Me.DBPath
    End If
    If strargDBFileName = "" Then
        '引数でDBファイル名が指定されていなかったらプロパティから引っ張る
        strargDBFileName = Me.DBFileName
    End If
    Dim fsoCreateConnectionString As FileSystemObject
    Set fsoCreateConnectionString = New FileSystemObject
    Dim strDBFileFullPath As String
    'DBファイルのフルパスを得る
    strDBFileFullPath = fsoCreateConnectionString.BuildPath(strargDBPath, strargDBFileName)
    'DBファイルの存在確認を行う、但しNoExistChekフラグが立ってたら無視する
    If (Not NoExistCheck) And (Not fsoCreateConnectionString.FileExists(strDBFileFullPath)) Then
        'ファイルが存在しなかった場合
        DebugMsgWithTime "CreateConnectionString: file not found"
        CreateConnectionString = Empty
        GoTo CloseAndExit
        Exit Function
    End If
    '引数で与えられたファイルがDBファイルかどうかを判定する
    If Not IsDBExtention(strDBFileFullPath) Then
        'DBファイルの拡張子ではなかった場合
        DebugMsgWithTime "CreateConnectionString: Not DB file Exetention."
        CreateConnectionString = Empty
        GoTo CloseAndExit
        Exit Function
    End If
    '拡張子により分岐し、それぞれの接続文字列を作成する
    Dim EnumValues As clsEnum
    Set EnumValues = CreateclsEnum
    Dim dicExtentions As Dictionary
    'Keyが拡張子、ValueがDB_file_exetension列挙体のDictionariを得る
    Set dicExtentions = EnumValues.DBFileExetension_Dic
    '拡張子により分岐
    Dim dicReplaceParm As Dictionary
    Set dicReplaceParm = New Dictionary
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = New clsSQLStringBuilder
    Select Case CLng(dicExtentions(LCase(fsoCreateConnectionString.GetExtensionName(strDBFileFullPath))))
    Case DB_file_exetension.accdb_dext
        'accdbの場合
        dicReplaceParm.Add 0, sqlBC.addQuote(strDBFileFullPath)
        CreateConnectionString = sqlBC.ReplaceParm(CONNECTION_STRING_ACCDB_0DBFullPath, dicReplaceParm)
    Case DB_file_exetension.xlam_dext, DB_file_exetension.xlsm_dext
        'xlsm(xlam)ファイルの場合
        dicReplaceParm.Add 0, sqlBC.addQuote(strDBFileFullPath)
        CreateConnectionString = sqlBC.ReplaceParm(CONNECTION_STRING_XLSM_0DBFullPath, dicReplaceParm)
    Case DB_file_exetension.xlsx_dext
        'xlsxの場合
        dicReplaceParm.Add 0, sqlBC.addQuote(strDBFileFullPath)
        CreateConnectionString = sqlBC.ReplaceParm(CONNECTION_STRING_XLSX_0DBFullPath, dicReplaceParm)
    Case DB_file_exetension.xlsb_dext
        'xlsbの場合
        dicReplaceParm.Add 0, sqlBC.addQuote(strDBFileFullPath)
        CreateConnectionString = sqlBC.ReplaceParm(CONNECTION_STRING_XLSB_0DBFullPath, dicReplaceParm)
    Case DB_file_exetension.xls_dext
        'xlsの場合
        dicReplaceParm.Add 0, sqlBC.addQuote(strDBFileFullPath)
        CreateConnectionString = sqlBC.ReplaceParm(CONNECTION_STRING_XLS_0DBFullPath, dicReplaceParm)
    Case DB_file_exetension.csv_dext
        'csvの場合
        'ここだけフルパスじゃないので注意
        dicReplaceParm.Add 0, sqlBC.addQuote(strargDBPath)
        CreateConnectionString = sqlBC.ReplaceParm(CONNECTION_STRING_CSV_0DBPathOnly, dicReplaceParm)
    Case Else
        '未定義だった場合
        DebugMsgWithTime "CreateConnectionString: Undefined extention"
        CreateConnectionString = Empty
        GoTo CloseAndExit
    End Select
    GoTo CloseAndExit
    Exit Function
CloseAndExit:
    Set sqlBC = Nothing
    Set fsoCreateConnectionString = Nothing
    Exit Function
End Function
'''DBへの接続を開始する
'''strargConnection 接続文字列をStringで渡す
Private Function OpenConnection(strargConnection As String) As Boolean
    If strargConnection = "" Then
        '接続文字列が指定されていなかったら即抜ける
        DebugMsgWithTime "OpenConnection Connection String Empty"
        OpenConnection = False
        Exit Function
    End If
    '接続の有無の確認方法を変更 外部で接続を閉じられても対応できるように 2021_12_19 Patacchi
    'プロパティのアクセサ側で対応
    If Me.Connected Then
        DebugMsgWithTime "OpenConnection already conected"
        OpenConnection = True
        Exit Function
    End If
    On Error GoTo ErrorCatch
    If adoConnection Is Nothing Then
        'クラスのメンバ変数が初期化されてなかったら初期化する
       Set adoConnection = New ADODB.Connection
    End If
    adoConnection.ConnectionString = Me.ConnectionString
    'ConnectionのCursorLocationをCliantにする。これをやらないと、RecorCountとかが取れなかったりする
    adoConnection.CursorLocation = adUseClient
    'ConnectionModeを指定する。デフォルトはReadOnly
    adoConnection.Mode = Me.ConnectMode
    adoConnection.Open
    If adoConnection.Errors.Count >= 1 Then
        GoTo ErrorCatch
        Exit Function
    End If
    isConnected = True
    OpenConnection = True
    Exit Function
ErrorCatch:
    If adoConnection.Errors.Count >= 1 Then
        '何らかのエラーが起きた
        Dim longErrorCount As Long
        For longErrorCount = 0 To adoConnection.Errors.Count - 1
            DebugMsgWithTime "OpenConnection code: " & adoConnection.Errors(longErrorCount).Number & " Description: " & adoConnection.Errors(longErrorCount).Description
        Next longErrorCount
        OpenConnection = False
        Exit Function
    End If
    DebugMsgWithTime "OpenConnection code: " & Err.Number & " Description: " & Err.Description
    OpenConnection = False
    Exit Function
End Function
'''Author Patacchi 2021_12_29
'''クラスのメンバのConnectionの接続を閉じる
Public Sub CloseClassConnection()
    If Me.Connected Then
        '接続されていたら閉じる
        adoConnection.Close
        isConnected = False
    End If
End Sub
'''DBディレクトリにカレントディレクトリを移動させる
'''ディレクトリが存在しなければ作成する
Private Function ChCurrentToDBDirectory(Optional strargDBPath As String = "") As Boolean
    On Error GoTo ErrorCatch
    'カレントディレクトリをDBディレクトリに移動する
    'カレントディレクトリの取得（UNCパス対応）
    'DataBaseディレクトリの存在有無確認"
    Dim strCurrentDir As String
    If strargDBPath <> "" Then
        '引数で指定されていたらそのディレクトリをDBディレクトリとする
        strCurrentDir = strargDBPath
    Else
        '指定されていない場合はConstで指定されているディレクトリにする
        strCurrentDir = constDatabasePath
    End If
'    If fso.FolderExists(constDatabasePath) <> True Then
    Dim fso As New scripting.FileSystemObject
    If fso.FolderExists(strCurrentDir) <> True Then
        'ディレクトリ存在しない場合作成しよ？
        MsgBox "データベースフォルダが無いため作成します。"
        MkDir strCurrentDir
    End If
    'データベースディレクトリに移動
    ChCurrentDirW (strCurrentDir)
    ChCurrentToDBDirectory = True
    Exit Function
ErrorCatch:
    DebugMsgWithTime "ChCurrenttoDB code: " & Err.Number & " Description: " & Err.Description
    If Err.Number = 76 Then
        'パスが見つかりません、mkdirでエラーが出た = ネットワークダメなんだね
        MsgBox "DBディレクトリ作成失敗（多分ネットワーク問題）処理を中断します"
    End If
    ChCurrentToDBDirectory = False
    Exit Function
End Function
Private Function InitialDBCreate(Optional strargDBFileName As String) As Boolean
    Dim isCollect As Boolean
    Dim strSQL  As String
    On Error GoTo ErrorCatch
    If strargDBFileName = "" Then
        'DBファイル名が引数で指定されていなかったらプロパティから引っ張る
        strargDBFileName = Me.DBFileName
        If strargDBFileName = "" Then
            'それでも空白だったら中止する
            DebugMsgWithTime "InitilaDBCreate File name empty"
            InitialDBCreate = False
            Exit Function
        End If
    End If
    'とりあえず空のDBファイルを作成する
    Dim adoxCat As ADOX.Catalog
    Set adoxCat = New ADOX.Catalog
    '接続文字列作成
    Call adoxCat.Create(Me.ConnectionString)
    Set adoxCat = Nothing
    InitialDBCreate = True
    Exit Function
ErrorCatch:
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
    End If
    Exit Function
End Function
'''Authour Daisuke Oota 2021_12_19
'''指定されたテーブル (データベースパス、データベースファイル名)に指定されたフィールドを追加する。既に存在していたら中断する
'''戻り値 Boolean 成功しているか、既に存在していたらTrue それ以外はFalse
'''オプションのstrargDataSizeはTEXT型かCHR型か
Public Function AppendField(strargTableName As String, strargFieldName As String, argaccdbType As ACCDB_Data_Type, Optional strargDataSize As String, _
                            Optional strargDBFileName As String = "", Optional strargDBPath As String = "") As Boolean
    'TablenameかFieldNameどちらかが空白だったら抜ける
    If strargTableName = "" Or strargFieldName = "" Then
        DebugMsgWithTime "AppendField: Tablename or Filename is empty"
        AppendField = False
        Exit Function
    End If
    '引数でDBpathおよびDBFilenameが指定されていなかったらプロパティから引っ張ってみる
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    'プロパティから引っ張っても空白だったら抜ける
    If strargDBPath = "" Or strargDBFileName = "" Then
        DebugMsgWithTime "AppenField: DBPath or DBFilename is empty"
        AppendField = False
        Exit Function
    End If
    '指定されたテーブルが存在するかチェックする
    Dim boolDBFileExists As Boolean
    boolDBFileExists = IsDBFileExist(strargDBFileName, strargDBPath)
    If Not boolDBFileExists Then
        'DBFileが存在しなかったら抜ける
        DebugMsgWithTime "AppendField: File or Path not exists"
        MsgBox "AppenField Path: " & strargDBPath & vbCrLf & " Filename: " & strargTableName & " is not exists"
        AppendField = False
        Exit Function
    End If
    '指定されたフィールド名が既に存在するかチェックする
    Dim IsFieldExists As Boolean
    IsFieldExists = Me.IsFieldExists(strargTableName, strargFieldName, strargDBFileName, strargDBPath)
    If IsFieldExists Then
        '既に指定のテーブルに指定のフィールドが存在していたら作業せずTrueを返して終了
        DebugMsgWithTime "AppendField:Table: " & strargTableName & " Field: " & strargFieldName & " is already exists."
        AppendField = True
        Exit Function
    End If
    Dim strSQL As String
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = CreateclsSQLStringBuilder
'    Set dicAppendField = New Dictionary
'    Dim clsEnumValue As clsEnum
'    Set clsEnumValue = CreateclsEnum
    'テーブル名にaddquoteすると失敗するみたい
'    dicAppendField.Add 0, sqlBc.addQuote(strargTableName)
'    dicAppendField.Add 0, strargTableName
    'フィールド名にaddquoteすると・・・・""込みのフィールド名になるのでNG！
'    dicAppendField.Add 1, sqlBc.addQuote(strargFieldName)
'    dicAppendField.Add 1, strargFieldName
'    dicAppendField.Add 2, clsEnumValue.AccdbDataType(argaccdbType)
    '置換パラメータ用Dictionary取得
    Dim dicAppendField As Dictionary
    Set dicAppendField = GetReplaceParmDic(strargTableName, strargFieldName, argaccdbType, strargDataSize)
    strSQL = sqlBC.ReplaceParm(SQL_APPEND_FIELD_0Tableneme_1fieldname_2DataType, dicAppendField)
    'ConnectModeのWriteフラグの調査・設定
    If Not Me.ConnectMode And adModeWrite Then
        Me.ConnectMode = Me.ConnectMode Or adModeWrite
    End If
    Dim isCollect As Boolean
    isCollect = Me.Do_SQL_with_NO_Transaction(strSQL, strargDBPath, strargDBFileName)
    'ConnectModeのWriteフラグを下す
    Me.ConnectMode = Me.ConnectMode And Not (adModeWrite)
    If Not isCollect Then
        'DoSqlで失敗してるようだったら抜ける
        DebugMsgWithTime "AppendField: fail when Do SQL"
        AppendField = False
        Exit Function
    End If
    '追加したフィールドが存在するかチェックする
    Dim isNewFieldExist As Boolean
    isNewFieldExist = Me.IsFieldExists(strargTableName, strargFieldName, strargDBFileName, strargDBPath)
    If isNewFieldExist Then
        '新しいフィールドが見つかった場合→成功
        DebugMsgWithTime "AppendField: " & strargFieldName & " on Table: " & strargTableName & " is success"
        AppendField = True
        Exit Function
    Else
        '新しいフィールドが見つからなかった場合→なぜ・・・？
        DebugMsgWithTime "AppendField: Fail AppendField: " & strargFieldName
        AppendField = False
        Exit Function
    End If
End Function
'''Author Patacchi 2021_12_26
'''テーブル名とフィールド名を引数として指定されたフィールドを削除する
'''戻り値 Boolean 削除成功したらTrue それ以外はFalseを返す
Public Function DeleteField(strargTableName As String, strargFieldName As String, _
                            Optional strargDBFileName As String = "", Optional strargDBPath As String = "") As Boolean
    'TablenameかFieldNameどちらかが空白だったら抜ける
    If strargTableName = "" Or strargFieldName = "" Then
        DebugMsgWithTime "DeleteField: Tablename or Filename is empty"
        DeleteField = False
        Exit Function
    End If
    '引数でDBpathおよびDBFilenameが指定されていなかったらプロパティから引っ張ってみる
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    'プロパティから引っ張っても空白だったら抜ける
    If strargDBPath = "" Or strargDBFileName = "" Then
        DebugMsgWithTime "DeleteField: DBPath or DBFilename is empty"
        DeleteField = False
        Exit Function
    End If
    '指定されたテーブルが存在するかチェックする
    Dim boolDBFileExists As Boolean
    boolDBFileExists = IsDBFileExist(strargDBFileName, strargDBPath)
    If Not boolDBFileExists Then
        'DBFileが存在しなかったら抜ける
        DebugMsgWithTime "DeleteField: File or Path not exists"
        MsgBox "AppenField Path: " & strargDBPath & vbCrLf & " Filename: " & strargTableName & " is not exists"
        DeleteField = False
        Exit Function
    End If
    '指定されたフィールド名が既に存在するかチェックする
    Dim IsFieldExists As Boolean
    IsFieldExists = Me.IsFieldExists(strargTableName, strargFieldName, strargDBFileName, strargDBPath)
    If Not IsFieldExists Then
        '既に指定のテーブルに指定のフィールドが存在していなかったら作業せずに終了
        DebugMsgWithTime "DeleteField:Table: " & strargTableName & " Field: " & strargFieldName & " is already NOT exists."
        DeleteField = False
        Exit Function
    End If
    Dim strSQL As String
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = CreateclsSQLStringBuilder
    Dim dicDeleteField As Dictionary
'    Set dicDeleteField = New Dictionary
'    dicDeleteField.Add 0, strargTableName
'    dicDeleteField.Add 1, strargFieldName
    Set dicDeleteField = GetReplaceParmDic(strargTableName, strargFieldName)
    strSQL = sqlBC.ReplaceParm(SQL_DELETE_FIELD_0Tablename_1Fieldname, dicDeleteField)
    'ConnectModeのWriteフラグの調査・設定
    If Not Me.ConnectMode And adModeWrite Then
        Me.ConnectMode = Me.ConnectMode Or adModeWrite
    End If
    Dim isCollect As Boolean
    isCollect = Me.Do_SQL_with_NO_Transaction(strSQL, strargDBPath, strargDBFileName)
    'Writeフラグを下す
    Me.ConnectMode = Me.ConnectMode And Not adModeWrite
    If Not isCollect Then
        'DoSqlで失敗してるようだったら抜ける
        DebugMsgWithTime "DeleteField: fail when Do SQL"
        DeleteField = False
        Exit Function
    End If
    '追加したフィールドが存在するかチェックする
    Dim isNewFieldExist As Boolean
    isNewFieldExist = Me.IsFieldExists(strargTableName, strargFieldName, strargDBFileName, strargDBPath)
    If Not isNewFieldExist Then
        '新しいフィールドが見つからなかった場合→成功
        DebugMsgWithTime "DeleteField: " & strargFieldName & " on Table: " & strargTableName & " is success"
        DeleteField = True
        Exit Function
    Else
        '新しいフィールドが見つかった場合→なぜ・・・？
        DebugMsgWithTime "DeleteField: Fail DeleteField: " & strargFieldName
        DeleteField = False
        Exit Function
    End If
End Function
'''Author Patacchi 2021_12_26
'''テーブル名とフィールド名、データ型を引数にとり、フィールドのデータ型を変更する
'''戻り値 Boolean 成功したらTrue それ以外はFalse
Public Function ChangeDataType(strargTableName As String, strargFieldName As String, argaccdbType As ACCDB_Data_Type, Optional strargDataSize As String, _
                            Optional strargDBFileName As String = "", Optional strargDBPath As String = "") As Boolean
    'TablenameかFieldNameどちらかが空白だったら抜ける
    If strargTableName = "" Or strargFieldName = "" Then
        DebugMsgWithTime "ChangeDataType: Tablename or Filename is empty"
        ChangeDataType = False
        Exit Function
    End If
    '引数でDBpathおよびDBFilenameが指定されていなかったらプロパティから引っ張ってみる
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    'プロパティから引っ張っても空白だったら抜ける
    If strargDBPath = "" Or strargDBFileName = "" Then
        DebugMsgWithTime "ChangeDataType: DBPath or DBFilename is empty"
        ChangeDataType = False
        Exit Function
    End If
    '指定されたテーブルが存在するかチェックする
    Dim boolDBFileExists As Boolean
    boolDBFileExists = IsDBFileExist(strargDBFileName, strargDBPath)
    If Not boolDBFileExists Then
        'DBFileが存在しなかったら抜ける
        DebugMsgWithTime "ChangeDataType: File or Path not exists"
        MsgBox "AppenField Path: " & strargDBPath & vbCrLf & " Filename: " & strargTableName & " is not exists"
        ChangeDataType = False
        Exit Function
    End If
    '指定されたフィールド名が存在するかチェックする
    Dim IsFieldExists As Boolean
    IsFieldExists = Me.IsFieldExists(strargTableName, strargFieldName, strargDBFileName, strargDBPath)
    If Not IsFieldExists Then
        '指定のテーブルに指定のフィールドが存在ていなかったら作業せずFalseを返して終了
        DebugMsgWithTime "ChangeDataType:Table: " & strargTableName & " Field: " & strargFieldName & " not found."
        ChangeDataType = False
        Exit Function
    End If
    Dim strSQL As String
    Dim sqlBC As clsSQLStringBuilder
    Set sqlBC = CreateclsSQLStringBuilder
    Dim dicChangeDataType As Dictionary
'    Set dicChangeDataType = New Dictionary
'    Dim clsEnumValue As clsEnum
'    Set clsEnumValue = CreateclsEnum
'    dicChangeDataType.Add 0, strargTableName
'    dicChangeDataType.Add 1, strargFieldName
'    dicChangeDataType.Add 2, clsEnumValue.AccdbDataType(argaccdbType)
    '置換用のパラメータのDictionaryを取得する
    Set dicChangeDataType = GetReplaceParmDic(strargTableName, strargFieldName, argaccdbType, strargDataSize)
    strSQL = sqlBC.ReplaceParm(SQL_CHANGE_DATATYPE_0Tablename_1Fieldname_2DataType, dicChangeDataType)
    'ConnectModeのWriteフラグの調査・設定
    If Not Me.ConnectMode And adModeWrite Then
        Me.ConnectMode = Me.ConnectMode Or adModeWrite
    End If
    Dim isCollect As Boolean
    isCollect = Me.Do_SQL_with_NO_Transaction(strSQL, strargDBPath, strargDBFileName)
    'wireフラグを下す
    Me.ConnectMode = Me.ConnectMode And Not adModeWrite
    If Not isCollect Then
        'DoSqlで失敗してるようだったら抜ける
        DebugMsgWithTime "ChangeDataType: fail when Do SQL"
        ChangeDataType = False
    Else
        ChangeDataType = True
    End If
    Set sqlBC = Nothing
End Function
'''Author Patacchi 2021_12_26
'''テーブル名、フィールド名、(データ型(サイズ))を引数に取り、置換パラメータ用のDictionaryを返す
'''戻り値 Dictionary 置換パラメータ用、Keyはインデックス番号
'''パラメータ並び順は、0_Tablename 1_FieldName 2_データタイプ（オプション込み）となる
'''strargDataSizeは、TEXT型、DECIMAL型 CHAR型、BINARY型のみで有効になる
'''DECIMAL型は、(全体の桁数(,少数以下の桁数))という指定になる
Private Function GetReplaceParmDic(strargTableName As String, strargFieldName As String, Optional argaccdbType As ACCDB_Data_Type = -1, Optional strargDataSize As String) As Dictionary
    If strargTableName = "" Or strargFieldName = "" Then
        'テーブル名とフィールド名は必須項目なので、なければ抜ける
        DebugMsgWithTime "GetReplaceParmDic: TableName or FieldName is empty"
        Exit Function
    End If
    Dim dicLocalParm As Dictionary
    Set dicLocalParm = New Dictionary
    Dim EnumValue As clsEnum
    Set EnumValue = CreateclsEnum
    dicLocalParm.Add 0, strargTableName
    dicLocalParm.Add 1, strargFieldName
    If Not argaccdbType = -1 Then
        'データ型が指定されている場合のみ実行
        If (argaccdbType = Text_typ Or argaccdbType = Char_ShortText_typ Or argaccdbType = Decimal_typ) And strargDataSize <> "" Then
            'サイズ指定可能な型で、かつサイズ指定されている場合
            dicLocalParm.Add 2, EnumValue.AccdbDataType(argaccdbType) & strargDataSize
        Else
            'データ型指定のない型、もしくはサイズ指定がされていない場合
            dicLocalParm.Add 2, EnumValue.AccdbDataType(argaccdbType)
        End If
    End If
    Set GetReplaceParmDic = dicLocalParm
End Function
'DBファイルの存在有無を確認する。初期テーブル作成は別プロシージャにて行う
Public Function IsDBFileExist(Optional strargDBFileName As String, Optional strargDBPath As String = "") As Boolean
    'DBFilenameおよびDBPathが指定されていなければプロパティから引っ張る（初期値を設定するはず）
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    'プロパティから引っ張ってもDBFilenameかDBPathが空白の場合は抜ける
    If strargDBFileName = "" Or strargDBPath = "" Then
        IsDBFileExist = False
        Exit Function
    End If
    'DBファイルの有無を確認する
    Dim fsoObj As FileSystemObject
    Set fsoObj = New FileSystemObject
    If Not fsoObj.FileExists(fsoObj.BuildPath(strargDBPath, strargDBFileName)) Then
        'PathとFilenameでフルパスを組み立て、ファイルの存在有無確認→無かった
        DebugMsgWithTime "IsDBFileExist: FullPath: " & fsoObj.BuildPath(strargDBPath, strargDBFileName) & " is not found"
        IsDBFileExist = False
        Exit Function
    End If
    '正常終了
    IsDBFileExist = True
End Function
'''Authour Daisuke Oota 2021_12_31
'''指定されたファイルがデータベースファイルかどうか判定する（拡張子ベース）
'''戻り値 Boolean DBFileならTrue それ以外はFalse
'''parms
'''strargFullPath   ファイルのフルパス
Public Function IsDBExtention(strargFilename_or_Fullpass As String) As Boolean
    Dim fsoExtension As FileSystemObject
    Set fsoExtension = New FileSystemObject
    Dim EnumValue As clsEnum
    Set EnumValue = CreateclsEnum
    Dim dicEnumExtention As Dictionary
    '拡張子EnumのDictionaryを取得する
    Set dicEnumExtention = EnumValue.DBFileExetension_Dic
    'キーに拡張子が入っているので、引数のフルパスより拡張子を取り出し一致するかチェック
    If dicEnumExtention.Exists(LCase(fsoExtension.GetExtensionName(strargFilename_or_Fullpass))) Then
        '存在していた場合
        IsDBExtention = True
    Else
        '存在していなかった場合
        IsDBExtention = False
    End If
    Exit Function
End Function
'''Authour Daisuke Oota 2021_12_19
'''指定されたデータベースパス、データベースファイル名で指定されたテーブルが存在するか確認する
'''戻り値 Boolean 存在してたらTrue それ以外はFalse
Public Function IsTableExists(strargTableName As String, Optional strargDBFileName As String = "", Optional strargDBPath As String = "") As Boolean
    'TableNameが空白だったら抜ける
    If strargTableName = "" Then
        DebugMsgWithTime "IsTalbeExists: TableName is empty"
        IsTableExists = False
        Exit Function
    End If
    'DBPathかDBFilenameが空白だったらプロパティから引っ張ってみる
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    'プロパティから引っ張ってもDBPath、DBFilenameが空白だったら抜ける
    If strargDBFileName = "" Or strargDBPath = "" Then
        DebugMsgWithTime "IsTableExists: Filename or Path is empty"
        IsTableExists = False
        Exit Function
    End If
    'DBFileの存在有無確認
    Dim isDBFileExists As Boolean
    isDBFileExists = Me.IsDBFileExist(strargDBFileName, strargDBPath)
    If Not isDBFileExists Then
        'ファイルが存在しなかったら抜ける
        DebugMsgWithTime "IsTableExists Path:  " & strargDBPath & vbCrLf & " Filename: " & strargDBFileName & " is not exists"
        IsTableExists = False
        Exit Function
    End If
    'テーブル一覧を取得
    'Connectionオブジェクトの作成、接続
    '接続文字列がクラスのプロパティと引数で与えられたのが一緒で、かつConnectedがTrueになってたら接続を再利用する
    Dim conTableExists As ADODB.Connection
    If Me.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName) Then
        '接続文字列が一致する場合は、connectedプロパティを調べて、接続されてなかったら接続してやる
        If Me.Connected Then
            '接続済みなら既存の接続を再利用してやる
            Set conTableExists = adoConnection
        Else
            '未接続ならクラスの接続メソッドを利用して接続する
            Dim isCollecte As Boolean
            isCollecte = DBReady(strargDBPath, strargDBFileName)
            If Not isCollecte Then
                '接続失敗したっぽい
                GoTo ErrorCatch
            End If
            Set conTableExists = adoConnection
        End If
    Else
        '接続文字列が一致しない場合は新規にConnectionオブジェクトを定義して接続してやる
        Set conTableExists = New ADODB.Connection
        conTableExists.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName)
        conTableExists.CursorLocation = adUseClient
        conTableExists.Open
    End If
    'OpenSchemaでスキーマ情報をRecordsetとして受け取る
    'adSchemaTables
    '制約列(条件にArrayで与える)
    'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE
    'Array(Empty,Empty,TableName,"TABLE")
    Dim rsTableExists As ADODB.Recordset
    Set rsTableExists = conTableExists.OpenSchema(adSchemaTables, Array(Empty, Empty, strargTableName, "TABLE"))
    If rsTableExists.RecordCount <= 0 Then
        'Recordcountが0以下ならないのが確定
        IsTableExists = False
        GoTo CloseAndExit
    Else
        '1以上なら存在するぽい
        IsTableExists = True
        GoTo CloseAndExit
    End If
CloseAndExit:
    Set rsTableExists = Nothing
    If Me.ConnectionString <> CreateConnectionString(strargDBPath, strargDBFileName) Then
        '接続文字列がクラスのものと一致しない場合は一時利用なので接続を明示的に閉じてやる
        conTableExists.Close
    End If
    Set conTableExists = Nothing
    Exit Function
ErrorCatch:
    Set rsTableExists = Nothing
    If Me.ConnectionString <> CreateConnectionString(strargDBPath, strargDBFileName) Then
        '接続文字列がクラスのものと一致しない場合は一時利用なので接続を明示的に閉じてやる
        conTableExists.Close
    End If
    Set conTableExists = Nothing
    IsTableExists = False
    DebugMsgWithTime "IsTableExists code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
'''Authour Daisuke Oota 2021_12_16
'''指定されたデータベースパス、データベースファイル名、テーブル名の中で指定されたフィールドが存在するかチェックする
'''戻り値 Boolean 存在してたらTrue それ以外はFalse
Public Function IsFieldExists(strargTableName As String, strargFieldName As String, _
                            Optional strargDBFileName As String = "", Optional strargDBPath As String = "") As Boolean
    '引数でDBPathとDBFilenameが指定されていなかったらプロパティから引っ張ってみる
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    'プロパティから引っ張ってもDBFilenameかDBPathどちらかが空白だったら抜ける
    If strargDBFileName = "" Or strargDBPath = "" Then
        DebugMsgWithTime "IsFieldExists: DBPath or DBFilename is empty"
        IsFieldExists = False
        Exit Function
    End If
    'テーブル名とフィールド名どちらかが指定されていなかったら抜ける
    If strargTableName = "" Or strargFieldName = "" Then
        DebugMsgWithTime "IsFieldExists: TableName or FieldName is empty"
        IsFieldExists = False
        Exit Function
    End If
    '指定されたDBFile（Path込み）が存在するかチェックする
    Dim isDBFileExists As Boolean
    isDBFileExists = IsDBFileExist(strargDBFileName, strargDBPath)
    If Not isDBFileExists Then
        'ファイルが存在しなかったら抜ける
        MsgBox "指定されたDBPath: " & strargDBPath & vbCrLf & " DBFilename: " & strargDBFileName & " は存在しないようです。"
        IsFieldExists = False
        Exit Function
    End If
    '指定されたテーブルが存在するかチェックする
    Dim isTableExist As Boolean
    isTableExist = Me.IsTableExists(strargTableName, strargDBFileName, strargDBPath)
    If Not isTableExist Then
        DebugMsgWithTime "IsFieldExists: Table not found: " & strargTableName
        IsFieldExists = False
        Exit Function
    End If
    'Connectionオブジェクトのopwnschemaメソッドでスキーマ情報を取得し、フィールドの有無をチェックする
    'Connecvtionオブジェクトの作成と接続を行う
    Dim conFieldExists As ADODB.Connection
    If Me.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName) Then
        '接続文字列が一致する場合は、connectedプロパティを調べて、接続されてなかったら接続してやる
        If Me.Connected Then
            '接続済みなら既存の接続を再利用してやる
            Set conFieldExists = adoConnection
        Else
            '未接続ならクラスの接続メソッドを利用して接続する
            Dim isCollecte As Boolean
            isCollecte = DBReady()
            If Not isCollecte Then
                '接続失敗したっぽい
                GoTo ErrorCatch
            End If
            Set conFieldExists = adoConnection
        End If
    Else
        '接続文字列が一致しない場合は新規にConnectionオブジェクトを定義して接続してやる
        Set conFieldExists = New ADODB.Connection
        conFieldExists.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName)
        conFieldExists.CursorLocation = adUseClient
        conFieldExists.Open
    End If
    'ConnectionオブジェクトのOpenSchemaメソッドを実行し、結果をレコードセットに受け取る
    'adSchemaColumns
    'TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,COLUMN_NAME
    'array(empty,empty,strargTableName,strargFieldName)
    Dim rsFieldExists As ADODB.Recordset
    Set rsFieldExists = conFieldExists.OpenSchema(adSchemaColumns, Array(Empty, Empty, strargTableName, strargFieldName))
    If rsFieldExists.RecordCount <= 0 Then
        'RecordCountが0以下なら指定のフィールドはない
        IsFieldExists = False
        GoTo CloseAndExit
    Else
        '1以上ならある
        IsFieldExists = True
        GoTo CloseAndExit
    End If
CloseAndExit:
    Set rsFieldExists = Nothing
    '接続文字列がクラスのプロパティと一致しない場合は一時利用なので明示的に接続を閉じる
    If Me.ConnectionString <> CreateConnectionString(strargDBPath, strargDBFileName) Then
        conFieldExists.Close
    End If
    Set conFieldExists = Nothing
    Exit Function
ErrorCatch:
    Set rsFieldExists = Nothing
    '接続文字列がクラスのプロパティと一致しない場合は一時利用なので明示的に接続を閉じる
    If Me.ConnectionString <> CreateConnectionString(strargDBPath, strargDBFileName) Then
        conFieldExists.Close
    End If
    Set conFieldExists = Nothing
    IsFieldExists = False
    DebugMsgWithTime "IsFielsExists code: " & Err.Number & " Description: " & Err.Description
    Exit Function
End Function
'''Author Patacchi 2022_01_16
'''指定されたDBファイルのテーブル一覧を取得する
'''戻り値 string配列 テーブル名を格納した1次元String配列
'''args
'''Optional strargDBFileName        DBFilename、未指定の場合はクラスのプロパティから引っ張る
'''Optional strargDBPath            DBPath、未指定の場合はクラスのプロパティから引っ張る
Public Function GetTableList(Optional strargDBFileName As String, Optional strargDBPath As String) As String()
    'DBPathかDBFilenameが空白だったらプロパティから引っ張ってみる
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    'プロパティから引っ張ってもDBPath、DBFilenameが空白だったら抜ける
    If strargDBFileName = "" Or strargDBPath = "" Then
        DebugMsgWithTime "GetTableList: Filename or Path is empty"
        GoTo CloseAndExit
        Exit Function
    End If
    'DBFileの存在有無確認
    Dim isDBFileExists As Boolean
    isDBFileExists = Me.IsDBFileExist(strargDBFileName, strargDBPath)
    If Not isDBFileExists Then
        'ファイルが存在しなかったら抜ける
        DebugMsgWithTime "GetTaleList Path:  " & strargDBPath & vbCrLf & " Filename: " & strargDBFileName & " is not exists"
        GoTo CloseAndExit
        Exit Function
    End If
    'Connectionオブジェクトの作成、接続
    '接続文字列がクラスのプロパティと引数で与えられたのが一緒で、かつConnectedがTrueになってたら接続を再利用する
    Dim conGetTableList As ADODB.Connection
    If Me.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName) Then
        '接続文字列が一致する場合は、connectedプロパティを調べて、接続されてなかったら接続してやる
        If Me.Connected Then
            '接続済みなら既存の接続を再利用してやる
            Set conGetTableList = adoConnection
        Else
            '未接続ならクラスの接続メソッドを利用して接続する
            Dim isCollecte As Boolean
            isCollecte = DBReady()
            If Not isCollecte Then
                '接続失敗したっぽい
                GoTo ErrorCatch
            End If
            Set conGetTableList = adoConnection
        End If
    Else
        '接続文字列が一致しない場合は新規にConnectionオブジェクトを定義して接続してやる
        Set conGetTableList = New ADODB.Connection
        conGetTableList.ConnectionString = CreateConnectionString(strargDBPath, strargDBFileName)
        conGetTableList.CursorLocation = adUseClient
        conGetTableList.Open
    End If
    'テーブル一覧を取得
    Dim adoxGetTableList As ADOX.Catalog
    Set adoxGetTableList = New ADOX.Catalog
    Set adoxGetTableList.ActiveConnection = conGetTableList
    Dim arrstrTables() As String
    Dim adoxTable As ADOX.Table
    Dim longTableCount As Long
    longTableCount = 0
    For Each adoxTable In adoxGetTableList.Tables
        If adoxTable.Type = "TABLE" Then
            ReDim Preserve arrstrTables(longTableCount)
            arrstrTables(longTableCount) = adoxTable.Name
            longTableCount = longTableCount + 1
        End If
    Next adoxTable
    GetTableList = arrstrTables
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "GetTableList code: " & Err.Number & " Description: " & Err.Description
    Dim arrError() As String
    GetTableList = arrError
    GoTo CloseAndExit
CloseAndExit:
    Set adoxTable = Nothing
    If Not adoxGetTableList Is Nothing Then
        Set adoxGetTableList = Nothing
    End If
'    If Not conGetTableList Is Nothing Then
'        conGetTableList.Close
'        Set conGetTableList = Nothing
'    End If
    Exit Function
End Function
'''一時ファイル置き場DBが無ければ作成する
'''args
'''optilnal strargDBPath        一時ファイルを作成する場所を指定する、未指定の場合はクラスのプロパティ
Private Sub CreateandCheckTempDB(Optional strargDBPath As String)
    On Error GoTo ErrorCatch
    If strargDBPath = "" Then
        'DBPathが指定されていなかったら初期値を設定してやる
        strargDBPath = constDatabasePath
    End If
    Dim fsoTempDB As FileSystemObject
    Set fsoTempDB = New FileSystemObject
    Dim strTempDBFullPath As String
    strTempDBFullPath = fsoTempDB.BuildPath(strargDBPath, PublicConst.TEMP_DB_FILENAME)
    If Not fsoTempDB.FileExists(strTempDBFullPath) Then
        '一時テーブル格納DBファイルがない場合
        '作成する
        Dim adoxTempDB As ADOX.Catalog
        Set adoxTempDB = New ADOX.Catalog
        adoxTempDB.Create (CreateConnectionString(strargDBPath, PublicConst.TEMP_DB_FILENAME, True))
        Set adoxTempDB = Nothing
    End If
    GoTo CloseAndExit
    Exit Sub
ErrorCatch:
    DebugMsgWithTime "CreateTempDB code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Set adoxTempDB = Nothing
    Set fsoTempDB = Nothing
    Exit Sub
End Sub
'''指定のテーブルを消去する
'''戻り値 Bool  成功したらTrue、それ以外はFalse
'''args
'''optional strargDBFileName        'DBのファイル名、未指定の場合はクラスのプロパティを適用
'''optional strargDBPath            'DBのディレクトリ、未指定の場合はクラスのプロパティを適用
Public Function DropTable(strargTableName As String, Optional strargDBFileName As String, Optional strargDBPath As String) As Boolean
    If strargTableName = "" Then
        'テーブル名が空だったら抜ける
        DebugMsgWithTime "DropTable Table name empty"
        DropTable = False
        GoTo CloseAndExit
        Exit Function
    End If
    'DBFilenameとDBPathが指定されていなかったらクラスのプロパティから引っ張る
    If strargDBFileName = "" Then
        strargDBFileName = Me.DBFileName
    End If
    If strargDBPath = "" Then
        strargDBPath = Me.DBPath
    End If
    '指定のテーブルが存在するか確認
    If Not IsTableExists(strargTableName, strargDBFileName, strargDBPath) Then
        '既に指定のテーブルが存在しなかった場合
        DebugMsgWithTime "DropTable: table not exists: " & strargTableName
        DropTable = True
        GoTo CloseAndExit
        Exit Function
    End If
    '現在のConnectionModeでWriteフラグが立っているか確認する
    If Not Me.ConnectMode And adModeWrite Then
        'Writeフラグを立てる
        Me.ConnectMode = Me.ConnectMode Or adModeWrite
    End If
    'テーブル消去実行
    Dim isCollect As Boolean
     isCollect = Do_SQL_with_NO_Transaction("DROP TABLE " & strargTableName, strargDBPath, strargDBFileName)
    If isCollect Then
        '成功したとき
        DebugMsgWithTime "Drop Table: " & strargTableName & " Success"
        DropTable = True
        GoTo CloseAndExit
        Exit Function
    Else
        '失敗したとき
        DebugMsgWithTime "Drop Table: " & strargTableName & "Fail"
        DropTable = False
        GoTo CloseAndExit
        Exit Function
    End If
CloseAndExit:
    Exit Function
End Function
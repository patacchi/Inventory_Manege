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
Const constDBFILENAME               As String = "CAT_Find.accdb"           'CAT情報格納用DBファイル名
Const constDBPROVIDER_Pri           As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source="
Const Field_Initialdate             As String = "InitialInputDate"      '各テーブル共通、初回入力時刻
Const Field_Update                  As String = "UpdateDate"            '各テーブル共通、最終更新時刻
Const constLONGDEFAULTARRAYSIZE     As Long = 15000                     '結果格納用配列の初期サイズ（暫定）
Const constLONGEXPANDARRAYSIZE      As Long = 3000                      '格納配列が足りない場合の1回の拡張量
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
Private longAffected As Long                                            '影響を受けた行数を格納する（UpdateとかInsertとか）
'プロパティ
'DBPath
Property Get DBPath() As String
    If strDBPath = "" Then
        'メンバ変数が初期状態なら定数をセットしてやる
        strDBPath = constDatabasePath
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
        '未接続ならプロパティのDBPathとDBFilenameから接続文字列を生成してやる
        strConnection = CreateConnectionString(Me.DBPath, Me.DBFileName)
        ConnectionString = strConnection
        Exit Property
    End If
End Property
Property Let ConnectionString(ByVal strargConnection As String)
    '既に接続済みの場合は変更させない
    If isConnected Then
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
'読み取り専用？
Property Get RS() As ADODB.Recordset
    Set RS = rsMyRecord
End Property
'RA_Array
'読み取り専用
Property Get RS_Array() As Variant
    RS_Array = varMyArray
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
        If isConnected Then
            Call adoConnection.Close
        End If
        Set adoConnection = Nothing
    End If
    'Command
    Set adoCommand = Nothing
    'Patrameters
    Set adoParameters = Nothing
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
''接続文字列を作成する
Private Function CreateConnectionString(Optional ByVal strargDBPath As String, Optional ByVal strargDBFileName As String) As String
    If strargDBPath = "" Then
        '引数でDBパスが指定されていなかったらプロパティから引っ張る
        strargDBPath = Me.DBPath
    End If
    If strargDBFileName = "" Then
        '引数でDBファイル名が指定されていなかったらプロパティから引っ張る
        strargDBFileName = Me.DBFileName
    End If
    Dim strlocalConnection As String
    Dim sqlBc As clsSQLStringBuilder
    Set sqlBc = New clsSQLStringBuilder
    strlocalConnection = constDBPROVIDER_Pri
    strlocalConnection = strlocalConnection & sqlBc.addQuote(strargDBPath & "\" & strargDBFileName)
    CreateConnectionString = strlocalConnection
    Set sqlBc = Nothing
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
Private Sub CloseClassConnection()
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
    'ディレクトリ移動は大元でやるので、個別には行わない
'    'DBディレクトリへ移動
'    Dim isDBDir As Boolean
'    isDBDir = ChCurrentToDBDirectory(Me.DBPath)
'    If Not isDBDir Then
'        MsgBox "DBディレクトリ作成失敗（ネットワークが原因かも）処理を中断します。"
'        InitialDBCreate = False
'        Exit Function
'    End If
'    If Me.DBFileName = "" Or Me.DBPath = "" Then
'        '接続文字列かディレクトリかどちらかが空白なら抜ける（失敗）
'        DebugMsgWithTime "InitialDBCreate: DB File or Directory not found"
'        InitialDBCreate = False
'        Exit Function
'    End If
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
    Dim adoxCat As adox.Catalog
    Set adoxCat = New adox.Catalog
    '接続文字列作成
    Call adoxCat.Create(Me.ConnectionString)
    Set adoxCat = Nothing
'    '初期テーブル作成用SQL文作成(T_Kishu)
'    strSQL = ""
'    strSQL = "CREATE TABLE IF NOT EXISTS """ & Table_Kishu & """ (" & vbCrLf & """"
'    strSQL = strSQL & Kishu_Header & """ TEXT NOT NULL UNIQUE," & vbCrLf & """"
'    strSQL = strSQL & Kishu_KishuName & """ TEXT NOT NULL UNIQUE," & vbCrLf & """"
'    strSQL = strSQL & Kishu_KishuNickname & """ TEXT NOT NULL UNIQUE," & vbCrLf & """"
'    strSQL = strSQL & Kishu_TotalKeta & """ NUMERIC NOT NULL," & vbCrLf & """"
'    strSQL = strSQL & Kishu_RenbanKetasuu & """ NUMERIC NOT NULL," & vbCrLf & """"
'    strSQL = strSQL & Field_Initialdate & """ TEXT DEFAULT CURRENT_TIMESTAMP," & vbCrLf & """"
'    strSQL = strSQL & Field_Update & """ TEXT)"
'    isCollect = dbSQLite3.DoSQL_No_Transaction(strSQL)
'
'    'エイリアステーブル作成
'    Call CreateAliasTable
'    Set dbSQLite3 = Nothing
''    'テスト実装_SQL作成テスト
''    isCollect = CreateTable_by_KishuName("Test15")
''    If Not isCollect Then
''        InitialDBCreate = False
''    End If
'    '正常終了
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
Public Function AppendField(strargTableName As String, strargFieldName As String, argaccdbType As ACCDB_Data_Type, _
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
    Dim sqlBc As clsSQLStringBuilder
    Set sqlBc = CreateclsSQLStringBuilder
    Dim dicAppendField As Dictionary
    Set dicAppendField = New Dictionary
    Dim clsEnumValue As clsEnum
    Set clsEnumValue = CreateclsEnum
    dicAppendField.Add 0, strargTableName
    dicAppendField.Add 1, strargTableName
    dicAppendField.Add 2, clsEnumValue.AccdbDataType(argaccdbType)
    strSQL = sqlBc.ReplaceParm(SQL_APPEND_FIELD_0Tableneme_1fieldname_2DataType, dicAppendField)
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
            isCollecte = DBReady()
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
        GoTo CloseandExit
    Else
        '1以上なら存在するぽい
        IsTableExists = True
        GoTo CloseandExit
    End If
CloseandExit:
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
        GoTo CloseandExit
    Else
        '1以上ならある
        IsFieldExists = True
        GoTo CloseandExit
    End If
CloseandExit:
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
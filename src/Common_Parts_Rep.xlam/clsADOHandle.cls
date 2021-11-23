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
Const constDatabasePath             As String = "C:\Users\q3005sbe\AppData\Local\Rep\Inventory_Manege\bin\Inventory_DB"      'データベースディレクトリ
Const constDBFILENAME               As String = "CAT_Find.accdb"           'IE情報格納用DBファイル名
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
Private strSql As String
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
    '既に接続済みの場合は変更させない
    If isConnected Then
        Exit Property
    End If
    If strargDBPath <> "" Then
        strDBPath = strargDBPath
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
    '既に接続済みの場合は変更させない
    If isConnected Then
        Exit Property
    End If
    If strargDBFileName <> "" Then
        strDBFileName = strargDBFileName
    End If
End Property
'ConnectionString
Property Get ConnectionString() As String
    If strConnection = "" Then
        '未定義ならDBPathとDBFilenameより初期値を返す
        strConnection = CreateConnectionString(Me.DBPath, Me.DBFileName)
    End If
    ConnectionString = strConnection
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
    Connected = isConnected
End Property
'SQL
Property Get SQL() As String
    SQL = strSql
End Property
Property Let SQL(ByVal strargSQL As String)
    If Not strargSQL = "" Then
        strSql = strargSQL
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
    strSql = ""
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
    If Connected Then
        '既に接続済みなら何もしないで抜ける（でもTrueよ）
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
'DBファイルの存在有無を確認し、なければ初期テーブル作成を行う
Private Function IsDBFileExist(Optional strargDBFileName As String) As Boolean
    'DBファイルの有無を確認する
    Dim isCurrendDbDir As Boolean
    'ディレクトリは個別指定する可能性があるので、ここでは移動しない
'    isCurrendDbDir = ChCurrentToDBDirectory
'    If Not isCurrendDbDir Then
'        'DBディレクトリが存在しない（しかもmkdir処理をしているはず）なので、ネットワークがダメ
'        MsgBox "ネットワーク接続の問題により、処理を中止します。"
'        IsDBFileExist = False
'        Exit Function
'    End If
    Dim fsoObj As FileSystemObject
    Set fsoObj = New FileSystemObject
    Dim strlocalDBFileName As String
    If strargDBFileName <> "" Then
        '引数でDBファイル名が指定されていたらそちらをセットしてやる
        strlocalDBFileName = strargDBFileName
    Else
        '指定されていなければ定数をセットしてやる
        strlocalDBFileName = constDBFILENAME
    End If
    If Not fsoObj.FileExists(strlocalDBFileName) Then
        MsgBox "DBファイルが見つからないようなので新規作成します"
        InitialDBCreate (strlocalDBFileName)
        IsDBFileExist = True
    End If
    '正常終了
    IsDBFileExist = True
End Function
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
    Dim strSql  As String
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
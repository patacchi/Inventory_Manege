VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJsonIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' JSONファイルの読み込み、書き込み等を扱うクラス
''' ファイル名未指定の場合は、DEFAULT_JSON_FILENAME定数に指定したファイル名とする(General.json)
Option Explicit
'''--------------------------------------------------------------------------------------------------------------
'''定数定義
Private Const DEFAULT_JSON_FILENAME As String = "INVGeneral.json"                           'ファイル名のデフォルト、全般設定Json
Private Const DEFAULT_JSON_FILEDIRECTORY As String = "SettingJson"                          '設定ファイルを置くデフォルトのディレクトリ、Common_Partsのカレントディレクトリの下に作成
Private Const KEY_INV_CONST As String = "INV_CONST"                                         'INV定数最上位階層のキー名
'テーブル共通のキー 定数
Private Const KEY_TYPTABLECOMMON As String = "typTableCommon"                               '各テーブル共通項目の最上位キー
Private Const KEY_TABLECOMMON_INPUTDATE As String = "F_InputDate"                           '入力日時 yyyy/mm/ddTHH:MM:SS.SSS の共通23桁
'部品（手配コード）マスターテーブルのキー 定数
Private Const KEY_TYPPARTSMASTER As String = "typPartsMaster"                               'パーツマスター 最上位キー名
Private Const KEY_PARTS_MASTER_TABLENAME As String = "T_INV_M_Parts"                        '手配コードマスターのテーブル名
Private Const KEY_INV_TEHAI_ID As String = "F_INV_Tehai_ID"                                 '手配コードのID、各テーブルにはこの値を設定する
Private Const KEY_INV_TEHAI_CODE As String = "F_INV_Tehai_Code"                             '手配コード
Private Const KEY_INV_MANEGE_SECTON As String = "F_INV_Manege_Section"                      '管理課
Private Const KEY_INV_SYSTEM_LOCATION_NO As String = "F_INV_System_Location_No"             'システム側の棚番号
Private Const KEY_INV_KISHU As String = "F_INV_Kishu"                                       '機種 J7 とか JL とか
Private Const KEY_INV_STORE_CODE As String = "F_INV_Store_Code"                             '貯蔵記号 FA,BS,BL
Private Const KEY_INV_DELIVER_LOT As String = "F_INV_Deliver_Lot"                           '払い出しロット
Private Const KEY_INV_FILL_LOT As String = "F_INV_Fill_Lot"                                 '補充ロット
Private Const KEY_INV_LEAD_TIME As String = "F_INV_Lead_Time"                               'リードタイム
Private Const KEY_INV_ORDER_AMOUNT As String = "F_INV_Order_Amount"                         '発注数
Private Const KEY_INV_ORDER_REMAIN As String = "F_INV_Order_Remain"                         '発注残数
Private Const KEY_INV_STOCK_AMOUNT As String = "F_INV_Stock_Amount"                         '在庫数
Private Const KEY_INV_TANA_ID As String = "F_INV_Tana_ID"                                   '棚番マスターのID、実際の内容は棚マスターテーブルで定義する
Private Const KEY_INV_SYSTEM_NAME As String = "F_INV_System_Name"                           'システム側の品名、CASEとか
Private Const KEY_INV_SYSTEM_SPEC As String = "F_INV_System_Spec"                           'システム側の型格、IMONOとか
Private Const KEY_INV_STORE_UNIT As String = "F_INV_Sotre_Unit"                             '貯蔵単位　PとかSETとかの
Private Const KEY_INV_SYSTEM_DESCRIPTION As String = "F_INV_System_Description"             'システム側の自由記述欄
Private Const KEY_INV_LOCAL_DESCRIPTION As String = "F_INV_Local_Description"               '4701独自の詳細を記述したい場合に使用する
Private Const KEY_INV_MANEGE_SECTION_SUB As String = "F_INV_Manege_Section_Sub"             'システム側の管理課サブ
Private Const KEY_INV_LABEL_NAME_1 As String = "F_INV_Label_Name_1"                         'BINカードラベルの品名1行目 CASE
Private Const KEY_INV_LABEL_NAME_2 As String = "F_INV_Label_Name_2"                         'BINカードラベルの品名2行目、行数の少ないものの品名はこのフィールドの値を使う LF470コア
Private Const KEY_INV_LABEL_REMARK_1 As String = "F_INV_Label_Remark_1"                     'BINカードラベルの備考1行目 錆びやすいので注意
Private Const KEY_INV_LABEL_REMARK_2 As String = "F_INV_Label_Remark_2"                     'BINカードラベルの備考2行目 保管の際乾燥剤をパウチ付き袋に入る事
'棚番マスター
Private Const KEY_TYPTANAMASTER As String = "typTanaMaster"                                 '棚番マスターの最上位キー名
Private Const KEY_TANA_MASTER_TABLENAME As String = "T_INV_M_Tana"                          '棚番マスターのテーブル名
Private Const KEY_INV_TANA_LOCAL_TEXT As String = "F_INV_Tana_Local_Text"                   '表示用などローカルで使用する棚番名 K05G B01
Private Const KEY_INV_TANA_SYSTEM_TEXT As String = "F_INV_Tana_System_Text"                 'システム側の棚番
'ラベル出力関係キー 定数
Private Const KEY_TYPLABEL = "typLabelSetting"                                              'ラベル関係定数最上位キー名
Private Const KEY_INV_LABELTEMPTABLE_NAME As String = "LabelTempTableName"                  'ラベル一時テーブル テーブル名
Private Const KEY_INV_FORMSTARTTIME As String = "F_FormStartTime"                           'フォームスタートタイム
Private Const KEY_INV_FORMSTARTTIME_FRENDRY As String = "F_FormStartTime_FriendryName"      'フォームスタート時間 フレンドリーネーム
Private Const KEY_INV_INPUTDATE_FRENDRY As String = "F_InputDate_FriendryName"              '入力日時フレンドリーネーム
Private Const KEY_INV_ORDERNUM As String = "F_OrderNumber"                                  'オーダーNo
Private Const KEY_INV_SAVEPOINT_FRENDTY As String = "F_SavePoint_FriendryName"              'SavePointフレンドリーネーム
Private Const KEY_INV_SAVEPOINT As String = "F_SavePoint"                                   'SavePoint
Private Const KEY_INV_TEHAICODE_LENGTH As String = "F_Tehaicode_Length"                     '手配コード文字列長
'在庫情報シートに関する定数
'Excelファイル名は日付をシリアル値とした文字列を付加するので、毎回変動する
Private Const KEY_TYPSHZAIKO As String = "typSHZaiko"                                       '在庫情報シートの最上位キー名
Private Const KEY_INV_SH_ZAIKO_SHETNAME As String = "在庫情報"                              '在庫検索でダウンロードできるExcelファイルの在庫情報シート名
Private Const KEY_INV_SH_ZAIKO_TEHAI_CODE As String = "手配コード"                          '手配コード
Private Const KEY_INV_SH_ZAIKO_MANEGE_SECTON As String = "管理課記号"                       '管理課
Private Const KEY_INV_SH_ZAIKO_SYSTEM_TANA_NO As String = "棚番"                            'システム側の棚番号
Private Const KEY_INV_SH_ZAIKO_KISHU As String = "手配機種"                                 '機種名
Private Const KEY_INV_SH_ZAIKO_STORE_CODE As String = "貯蔵記号"                            '貯蔵記号
Private Const KEY_INV_SH_ZAIKO_DELIVER_LOT As String = "払出ロット"                         '払い出しロット
Private Const KEY_INV_SH_ZAIKO_FILL_LOT As String = "補充点数量"                            '補充ロット
Private Const KEY_INV_SH_ZAIKO_LEAD_TIME As String = "リードタイム"                         'リードタイム
Private Const KEY_INV_SH_ZAIKO_ORDER_AMOUNT As String = "発注数"                            '発注数
Private Const KEY_INV_SH_ZAIKO_ORDER_REMAIN As String = "発注残"                            '発注残数
Private Const KEY_INV_SH_ZAIKO_STOCK_AMOUNT As String = "在庫数量"                          '在庫数
Private Const KEY_INV_SH_ZAIKO_SYSTEM_NAME As String = "品名記号"                           'システム側の品名、CASEとか
Private Const KEY_INV_SH_ZAIKO_SYSTEM_SPEC As String = "型格記事"                           'システム側の型格、IMONOとか
Private Const KEY_INV_SH_ZAIKO_STORE_UNIT As String = "単位"                                '貯蔵単位　PとかSETとかの
Private Const KEY_INV_SH_ZAIKO_SYSTEM_DESCRIPTION As String = "在庫自由記述"                'システム側の自由記述欄
Private Const KEY_INV_SH_ZAIKO_MANEGE_SECTION_SUB As String = "管理課サブ"                  'システム側の管理課サブ
Private Const KEY_INV_SH_ZAIKO_TANA_TEXT As String = "ロケーション"                         '棚番号名前、DBには棚マスターテーブルから引っ張ったIDをセットする
'棚卸CSVに関する定数
Private Const KEY_TYPCSVTANA As String = "typCSVTana"                                       '棚卸CSVの最上位キー
Private Const KEY_INV_CSV_TANA_TABLENAME As String = "T_INV_CSV"                            '棚卸CSVファイルを格納するテーブル名
Private Const KEY_INV_CSV_TANA_ID As String = "F_CSV_ID"                                    'CSVファイルのID(重複有り、そのままでは使用できない)
Private Const KEY_INV_CSV_TANA_ENDDAY As String = "棚卸締切日"
Private Const KEY_INV_CSV_TANA_NO As String = "ＮＯ．"
'Private Const KEY_INV_CSV_TANA_MANEGE_SECTION = KEY_INV_SH_ZAIKO_MANEGE_SECTON              '値が同じ定数定義は不要(構造体で共有指定すれば良いため)
Private Const KEY_INV_CSV_TANA_MANEGE_SECTION_SUB = "サブコード"
'Private Const KEY_INV_CSV_TANA_TEHAI_CODE = KEY_INV_SH_ZAIKO_TEHAI_CODE                     '値が同じ定数定義の別名は不要
Private Const KEY_INV_CSV_TANA_SYSTEM_NAME As String = "品名"
Private Const KEY_INV_CSV_TANA_SYSTEM_SPEC As String = "型格"
'private Const KEY_INV_CSV_TANA_SYSTEM_TANA_NO = F_SH_ZAIKO_SYSTEM_TANA_NO                   '値が同じ定数定義の別名は不要
'private Const KEY_INV_CSV_TANA_LOCATION_TEXT = F_SH_ZAIKO_TANA_TEXT                         '値が同じ定数定義の別名は不要
'Private Const KEY_INV_CSV_TANA_STORE_CODE = F_SH_ZAIKO_STORE_CODE                           '値が同じ定数定義の別名は不要
Private Const KEY_INV_CSV_TANA_STOCK_AMOUNT As String = "在庫数"
Private Const KEY_INV_CSV_TANA_AVAILABLE_AMOUNT As String = "現品残"
Private Const KEY_INV_CSV_TANA_STATUS As String = "F_CSV_Status"                            'チェック状態フラグを管理するLong、独自フィールド
Private Const KEY_INV_CSV_TANA_BIN_AMOUNT As String = "F_CSV_BIN_Amount"                    'BINカード残数を記録するフィールド、独自フィールド
'''--------------------------------------------------------------------------------------------------------------
'''メンバ変数定義
Private strFilePath As String                                       'ファイルパスを格納するメンバ変数
Private fsoJsonFile As FileSystemObject                             'Jsonファイルのファイルシステムオブジェクト
Private typLabelSetting_ As typLabelSetting                         '印刷出力に関する定数定義の構造体
Private dicJsonParsed As Dictionary                                 'Jsonファイルをパースシテ結果が格納されるDictionary
'''--------------------------------------------------------------------------------------------------------------
'''プロパティ定義
''' Autthor Daisuke Oota 2023_03_23
''' JSONファイルフルパスを取得・設定するプロパティ
''' Filepath取得
''' 引数
'Property Get FilePath(Optional IsReload As Boolean = False) As String
Property Get FilePath() As String
    If strFilePath = "" Then
        'メンバ変数のファイルパスが空だったら標準設定を適用する
        Dim fsoDefaultFilepath As FileSystemObject
        Set fsoDefaultFilepath = New FileSystemObject
        strFilePath = fsoDefaultFilepath.BuildPath(Application.ThisWorkbook.Path & "\" & DEFAULT_JSON_FILEDIRECTORY, DEFAULT_JSON_FILENAME)
        FilePath = strFilePath
        Exit Property
    Else
        FilePath = strFilePath
        Exit Property
    End If
End Property
''' FilePath 設定
Property Let FilePath(strargFilePath As String)
    If strargFilePath = "" Then
        '空だったら抜ける
        DebugMsgWithTime "Property Let FilePath:Arg is Empty"
        Exit Property
    End If
    strFilePath = strargFilePath
End Property
''' LabelSetting構造体
Property Get LabelSetting() As typLabelSetting
    'ファイルからメンバ変数のDictionaryにパースした結果をセットする
    On Error GoTo ErrorCatch
    Call SetJsonDicbyFileName
    '構造体のメンバーに値をセットしていく
    typLabelSetting_.F_FormStartTime_FriendryName = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_FORMSTARTTIME_FRENDRY)
    typLabelSetting_.F_InputDate_FriendryName = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_INPUTDATE_FRENDRY)
    typLabelSetting_.F_OrderNumber = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_ORDERNUM)
    typLabelSetting_.F_FormStartTime = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_FORMSTARTTIME)
    typLabelSetting_.F_SavePoint_FriendryName = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_SAVEPOINT_FRENDTY)
    typLabelSetting_.F_SavePoint = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_SAVEPOINT)
    typLabelSetting_.F_Tehaicode_Length = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_TEHAICODE_LENGTH)
    typLabelSetting_.LabelTempTableName = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_LABELTEMPTABLE_NAME)
    typLabelSetting_.F_INV_Tana_Local_Text = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_TANA_LOCAL_TEXT)
    typLabelSetting_.F_INV_Tehai_Code = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_TEHAI_CODE)
    typLabelSetting_.F_INV_Store_Code = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_STORE_CODE)
    typLabelSetting_.F_INV_Kishu = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_KISHU)
    typLabelSetting_.F_INV_Label_Name_1 = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_LABEL_NAME_1)
    typLabelSetting_.F_INV_Label_Name_2 = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_LABEL_NAME_2)
    typLabelSetting_.F_INV_Label_Remark_1 = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_LABEL_REMARK_1)
    typLabelSetting_.F_INV_Label_Remark_2 = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_INV_LABEL_REMARK_2)
    typLabelSetting_.F_InputDate = dicJsonParsed(KEY_INV_CONST)(KEY_TYPLABEL)(KEY_TABLECOMMON_INPUTDATE)
    '取得した結果を返す
    LabelSetting = typLabelSetting_
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "LabelSetting code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
    Exit Property
CloseAndExit:
    Exit Property
End Property
'''--------------------------------------------------------------------------------------------------------------
''' プロシージャ定義
''' 指定したファイルパスのファイルをパースし、メンバ変数のDictionaryにセットする
Private Sub SetJsonDicbyFileName(Optional strargFilePath As String)
    On Error GoTo ErrorCatch
    Dim strTargetFilePath As String
    If strargFilePath = "" Then
        '引数でファイルパスが指定されなかった場合
        'プロパティから引っ張る
        strTargetFilePath = Me.FilePath
    Else
        '引数で指定された場合はその値をセットする
        strTargetFilePath = strargFilePath
    End If
    If fsoJsonFile Is Nothing Then
        'メンバ変数が初期化されていなかったらここで初期化
        Set fsoJsonFile = New FileSystemObject
    End If
    If Not fsoJsonFile.FileExists(strTargetFilePath) Then
        'ファイルが見つからなかった場合
        DebugMsgWithTime "SetJsonDicbyFileName: File Not found"
        GoTo CloseAndExit
        Exit Sub
    End If
    Dim adoStream As ADODB.Stream
'    Set tsJsonFile = fsoJsonFile.OpenTextFile(strTargetFilePath, ForReading, False, TristateFalse)
    Set adoStream = New ADODB.Stream
    'adoStreamのオプションをセットしていく
    'FileSystemObjectのOpenTextメソッドだとUTF-8ファイルが上手く扱えないので、ADODB.STREAMにて処理する
    adoStream.Charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile strTargetFilePath
    'Streamの結果をテキストとして読み込む
    Dim strJsonRaw As String
    strJsonRaw = adoStream.ReadText
    adoStream.Close
    'テキストからパースした結果をメンバ変数のDictionaryにセットする
    Set dicJsonParsed = JsonConverter.ParseJson(strJsonRaw)
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "SetJsonDicbyFileName code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
    Exit Sub
CloseAndExit:
    If Not adoStream Is Nothing Then
        If adoStream.State And adStateOpen Then
            '接続さてているようならCloseする
            adoStream.Close
        End If
    End If
    Exit Sub
End Sub
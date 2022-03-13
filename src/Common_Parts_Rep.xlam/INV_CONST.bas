Attribute VB_Name = "INV_CONST"
''''在庫管理関係の定数を定義する
Option Explicit
'在庫情報DBに関する定数
Public Const INV_DB_FILENAME As String = "INV_Manege.accdb"                     '在庫情報のDBファイル名
Public Const T_INV_TEMP As String = "T_INV_Temp"                                'INVDBの一時テーブル名
Public Const T_INV_SELECT_TEMP As String = "T_INV_Select_Temp"                  'Selectした結果を格納するテーブル名、一旦テーブルに格納しないと
                                                                                '更新可能なクエリが・・・とか言われるため
Public Const INV_DOC_LABEL_MAILMERGE As String = "INV_Label_Mailmerge_Local.docm"           'ラベル差し込み印刷のフィールド設定済みWordDocument名
Public Const INV_DOC_LABEL_PLANE As String = "INV_Label_MailmergePlain_Local.docm"          'ラベル差し込みの出力用空白Document名
Public Const INV_DOC_LABEL_GENPIN_SMALL As String = "INV_Genpin_Small_Local.docx"           '現品票(小)の差し込み印刷テンプレート
Public Const INV_DOC_LABEL_SPECSHEET_Small As String = "INV_SpecSheet_Small_Local.docx"     'フル記載(ラベルと同じ+オーダーNo) 小 テンプレート
'部品（手配コード）マスターテーブルの定数
Public Const T_INV_M_Parts As String = "T_INV_M_Parts"                          '手配コードマスターのテーブル名
Public Const F_INV_TEHAI_ID As String = "F_INV_Tehai_ID"                        '手配コードのID、各テーブルにはこの値を設定する
Public Const F_INV_TEHAI_TEXT As String = "F_INV_Tehai_Code"                    '手配コード
Public Const F_INV_MANEGE_SECTON As String = "F_INV_Manege_Section"             '管理課
Public Const F_INV_SYSTEM_LOCATION_NO As String = "F_INV_System_Location_No"    'システム側の棚番号
Public Const F_INV_KISHU As String = "F_INV_Kishu"                              '機種名
Public Const F_INV_STORE_CODE As String = "F_INV_Store_Code"                    '貯蔵記号
Public Const F_INV_DELIVER_LOT As String = "F_INV_Deliver_Lot"                  '払い出しロット
Public Const F_INV_FILL_LOT As String = "F_INV_Fill_Lot"                        '補充ロット
Public Const F_INV_LEAD_TIME As String = "F_INV_Lead_Time"                      'リードタイム
Public Const F_INV_ORDER_AMOUNT As String = "F_INV_Order_Amount"                '発注数
Public Const F_INV_ORDER_REMAIN As String = "F_INV_Order_Remain"                '発注残数
Public Const F_INV_STOCK_AMOUNT As String = "F_INV_Stock_Amount"                '在庫数
Public Const F_INV_TANA_ID As String = "F_INV_Tana_ID"                          '棚番マスターのID、実際の内容は棚マスターテーブルで定義する
Public Const F_INV_SYSTEM_NAME As String = "F_INV_System_Name"                  'システム側の品名、CASEとか
Public Const F_INV_SYSTEM_SPEC As String = "F_INV_System_Spec"                  'システム側の型格、IMONOとか
Public Const F_INV_STORE_UNIT As String = "F_INV_Sotre_Unit"                    '貯蔵単位　PとかSETとかの
Public Const F_INV_SYSTEM_DESCRIPTION As String = "F_INV_System_Description"    'システム側の自由記述欄
Public Const F_INV_LOCAL_DESCRIPTION As String = "F_INV_Local_Description"      '4701独自の詳細を記述したい場合に使用する
Public Const F_INV_MANEGE_SECTION_SUB As String = "F_INV_Manege_Section_Sub"    'システム側の管理課サブ
Public Const F_INV_LABEL_NAME_1 As String = "F_INV_Label_Name_1"                'BINカードラベルの品名1行目 CASE
Public Const F_INV_LABEL_NAME_2 As String = "F_INV_Label_Name_2"                'BINカードラベルの品名2行目、行数の少ないものの品名はこのフィールドの値を使う LF470コア
Public Const F_INV_LABEL_REMARK_1 As String = "F_INV_Label_Remark_1"            'BINカードラベルの備考1行目 錆びやすいので注意
Public Const F_INV_LABEL_REMARK_2 As String = "F_INV_Label_Remark_2"            'BINカードラベルの備考2行目 保管の際乾燥剤をパウチ付き袋に入る事
'手配コードマスターのEnum定義
Public Enum Enum_INV_M_Parts
    Table_Name_IMPrt = 1
    F_Tehai_ID_IMPrt = 2
    F_Tehai_Code_IMPrt = 3
    F_Manege_Section_IMPrt = 4
    F_System_TanaNo_IMPrt = 5
    F_Kishu_IMPrt = 6
    F_Store_Code_IMPrt = 7
    F_Deliver_Lot_IMPrt = 8
    F_Fill_Lot_IMPrt = 9
    F_Lead_Time_IMPrt = 10
    F_Order_Amount_IMPrt = 11
    F_Order_Remain_IMPrt = 12
    F_Stock_Amount_IMPrt = 13
    F_Tana_ID_IMPrt = 14
    F_System_Name_IMPrt = 15
    F_System_Spec_IMPrt = 16
    F_Store_Unit_IMPrt = 17
    F_System_Description_IMPrt = 18
    F_Local_Description_IMPrt = 19
    F_Manege_Section_Sub_IMPrt = 20
    F_InputDate_IMPrt = 21
    'BINカード、表示関係でローカル限定
    F_Label_Name_1_IMPrt = 22
    F_Label_Name_2_IMPrt = 23
    F_Label_Remark_1_IMPrt = 24
    F_Label_Remark_2_IMPrt = 25
End Enum
'在庫情報シートに関する定数
'Excelファイル名は日付をシリアル値とした文字列を付加するので、毎回変動する
Public Const INV_SH_ZAIKO_NAME As String = "在庫情報"                       '在庫検索でダウンロードできるExcelファイルの在庫情報シート名
Public Const F_SH_ZAIKO_TEHAI_TEXT As String = "手配コード"                 '手配コード
Public Const F_SH_ZAIKO_MANEGE_SECTON As String = "管理課記号"              '管理課
Public Const F_SH_ZAIKO_SYSTEM_TANA_NO As String = "棚番"                   'システム側の棚番号
Public Const F_SH_ZAIKO_KISHU As String = "手配機種"                        '機種名
Public Const F_SH_ZAIKO_STORE_CODE As String = "貯蔵記号"                   '貯蔵記号
Public Const F_SH_ZAIKO_DELIVER_LOT As String = "払出ロット"                '払い出しロット
Public Const F_SH_ZAIKO_FILL_LOT As String = "補充点数量"                   '補充ロット
Public Const F_SH_ZAIKO_LEAD_TIME As String = "リードタイム"                'リードタイム
Public Const F_SH_ZAIKO_ORDER_AMOUNT As String = "発注数"                   '発注数
Public Const F_SH_ZAIKO_ORDER_REMAIN As String = "発注残"                   '発注残数
Public Const F_SH_ZAIKO_STOCK_AMOUNT As String = "在庫数量"                 '在庫数
Public Const F_SH_ZAIKO_SYSTEM_NAME As String = "品名記号"                  'システム側の品名、CASEとか
Public Const F_SH_ZAIKO_SYSTEM_SPEC As String = "型格記事"                  'システム側の型格、IMONOとか
Public Const F_SH_ZAIKO_STORE_UNIT As String = "単位"                       '貯蔵単位　PとかSETとかの
Public Const F_SH_ZAIKO_SYSTEM_DESCRIPTION As String = "在庫自由記述"       'システム側の自由記述欄
Public Const F_SH_ZAIKO_MANEGE_SECTION_SUB As String = "管理課サブ"         'システム側の管理課サブ
Public Const F_SH_ZAIKO_TANA_TEXT As String = "ロケーション"                '棚番号名前、DBには棚マスターテーブルから引っ張ったIDをセットする
'在庫情報シートのEnum
'大部分をINV_Master_PartsのEnumと共有する
'棚番のみは共有できないので、独自に数字を振る
Public Enum Enum_Sh_Zaiko
    F_Tehai_Code_ShZ = Enum_INV_M_Parts.F_Tehai_Code_IMPrt
    F_Manege_Section_ShZ = Enum_INV_M_Parts.F_Manege_Section_IMPrt
    F_System_TanaNO_ShZ = Enum_INV_M_Parts.F_System_TanaNo_IMPrt
    F_kishu_ShZ = Enum_INV_M_Parts.F_Kishu_IMPrt
    F_Store_Code_ShZ = Enum_INV_M_Parts.F_Store_Code_IMPrt
    F_Deliver_Lot_ShZ = Enum_INV_M_Parts.F_Deliver_Lot_IMPrt
    F_Fill_Lot_ShZ = Enum_INV_M_Parts.F_Fill_Lot_IMPrt
    F_Lead_Time_ShZ = Enum_INV_M_Parts.F_Lead_Time_IMPrt
    F_Order_Amount_ShZ = Enum_INV_M_Parts.F_Order_Amount_IMPrt
    F_Order_Remain_ShZ = Enum_INV_M_Parts.F_Order_Remain_IMPrt
    F_Stock_Amount_ShZ = Enum_INV_M_Parts.F_Stock_Amount_IMPrt
    F_System_Name_ShZ = Enum_INV_M_Parts.F_System_Name_IMPrt
    F_System_Spec_ShZ = Enum_INV_M_Parts.F_System_Spec_IMPrt
    F_Store_Unit_ShZ = Enum_INV_M_Parts.F_Store_Unit_IMPrt
    F_System_Description_ShZ = Enum_INV_M_Parts.F_System_Description_IMPrt
    F_Manege_Section_Sub_ShZ = Enum_INV_M_Parts.F_Manege_Section_Sub_IMPrt
    '棚番テキストのみこちらで独自に設定する100番台～
    F_Tana_Text_ShZ = 101
End Enum
'棚番マスター
Public Const T_INV_M_Tana As String = "T_INV_M_Tana"                            '棚番マスターのテーブル名
'フィールド名定数
Public Const F_INV_TANA_LOCAL_TEXT As String = "F_INV_Tana_Local_Text"              '表示用などローカルで使用する棚番名 K05G B01
Public Const F_INV_TANA_SYSTEM_TEXT As String = "F_INV_Tana_System_Text"            'システム側の棚番
Public Const F_INV_TANA_TIET_DELIVARY As String = "F_INV_TIET_Delivery"                  'TIET出庫の棚かどうか
'T_M_Tanaフィールド定義Enum
Public Enum Enum_INV_M_Tana
    '共通フィールド
    F_INV_TANA_ID_IMT = Enum_INV_M_Parts.F_Tana_ID_IMPrt
    'Tanaテーブルのみにあるのは100番台にする
    F_INV_Tana_Local_Text_IMT = 102
    F_INV_Tana_System_Text_IMT = 103
    F_INV_TIET_Delivary_IMT = 104
    F_InputDate_IMT = 105
End Enum
'在庫情報シートでUpdate掛ける際にTrim必要なフィールド名を定義
'_ntrm need trim
Public Enum Enum_SH_Zaiko_Need_Trim
    F_Manege_Section_ntrm = Enum_Sh_Zaiko.F_Manege_Section_ShZ
    F_Tehai_Code_ntrm = Enum_Sh_Zaiko.F_Tehai_Code_ShZ
    F_System_TanaNO_ntrm = Enum_Sh_Zaiko.F_System_TanaNO_ShZ
    F_Kishu_ntrm = Enum_Sh_Zaiko.F_kishu_ShZ
    F_Store_Code_ntrm = Enum_Sh_Zaiko.F_Store_Code_ShZ
    F_Tana_Text_ntrm = Enum_Sh_Zaiko.F_Tana_Text_ShZ
    F_System_Name_ntrm = Enum_Sh_Zaiko.F_System_Name_ShZ
    F_System_Spec_ntrm = Enum_Sh_Zaiko.F_System_Spec_ShZ
    F_Store_Unit_ntrm = Enum_Sh_Zaiko.F_Store_Unit_ShZ
    F_Manege_Section_sub_ntrm = Enum_Sh_Zaiko.F_Manege_Section_Sub_ShZ
    F_System_Description_ntrm = Enum_Sh_Zaiko.F_System_Description_ShZ
End Enum
'棚卸CSVファイル
'T_M_Partsとは別にデータを持つので、それほど気にしなくてもいいかも
'F_CSV_ID    （オートインクリメント）
'ＮＯ．         独自
'棚卸締切日     ローカル
'管理課記号     共通
'サブコード
'手配コード     共通
'品名
'型格
'棚番 共通
'ロケーション   共通
'貯蔵記号 共通
'在庫数
'現品残         ローカル限定
'DB格納テーブル名
Public Const T_INV_CSV As String = "T_INV_CSV"                                                      '棚卸CSVファイルを格納するテーブル名
Public Const F_INV_CSV_ID As String = "F_CSV_ID"
Public Const F_INV_CSV_ENDDAY As String = "棚卸締切日"
Public Const F_INV_CSV_NO As String = "ＮＯ．"
Public Const F_INV_CSV_MANEGE_SECTION = F_SH_ZAIKO_MANEGE_SECTON
Public Const F_INV_CSV_MANEGE_SECTION_SUB = "サブコード"
Public Const F_INV_CSV_TEHAI_TEXT = F_SH_ZAIKO_TEHAI_TEXT
Public Const F_INV_CSV_SYSTEM_NAME As String = "品名"
Public Const F_INV_CSV_SYSTEM_SPEC As String = "型格"
Public Const F_INV_CSV_SYSTEM_TANA_NO = F_SH_ZAIKO_SYSTEM_TANA_NO
Public Const F_INV_CSV_LOCATION_TEXT = F_SH_ZAIKO_TANA_TEXT
Public Const F_INV_CSV_STORE_CODE = F_SH_ZAIKO_STORE_CODE
Public Const F_INV_CSV_STOCK_AMOUNT As String = "在庫数"
Public Const F_INV_CSV_AVAILABLE_AMOUNT As String = "現品残"
Public Const F_INV_CSV_STATUS As String = "F_CSV_Status"                                            'チェック状態フラグを管理するLong、独自フィールド
Public Const F_INV_CSV_BIN_AMOUNT As String = "F_CSV_BIN_Amount"                                    'BINカード残数を記録するフィールド、独自フィールド
'CSV Enum Inv CSv file
'3個以外は共通なので共通(?)
Public Enum Enum_CSV_Tana_Field
    F_EndDay_ICS = 100
    F_CSV_No_ICS = 101
    F_Available_ICS = 102
    'ロケーションだけメンドウなので独自に番号振る
    F_Location_Text_ICS = 103
    F_Status_ICS = 104
    F_Bin_Amount_ICS = 105
    '以下は共通フィールド
    F_ManegeSection_ICS = Enum_INV_M_Parts.F_Manege_Section_IMPrt
    F_ManegeSection_Sub_ICS = Enum_INV_M_Parts.F_Manege_Section_Sub_IMPrt
    F_Tehai_Code_ICS = Enum_INV_M_Parts.F_Tehai_Code_IMPrt
    F_System_Name_ICS = Enum_INV_M_Parts.F_System_Name_IMPrt
    F_System_Spec_ICS = Enum_INV_M_Parts.F_System_Spec_IMPrt
    F_System_Tana_NO_ICS = Enum_INV_M_Parts.F_System_TanaNo_IMPrt
    F_Store_Code_ICS = Enum_INV_M_Parts.F_Store_Code_IMPrt
    F_Stock_Amount_ICS = Enum_INV_M_Parts.F_Stock_Amount_IMPrt
End Enum
'棚卸CSVファイルでTrim必要なフィールドの定数を列挙する Csv need TRiM _ctrm
Public Enum Enum_INV_CSV_Need_Trim
    F_EndDay_ctrm = Enum_CSV_Tana_Field.F_EndDay_ICS
    F_Manege_Section_ctrm = Enum_CSV_Tana_Field.F_ManegeSection_ICS
    F_Manege_Section_Sub_ctrm = Enum_CSV_Tana_Field.F_ManegeSection_Sub_ICS
    F_Tehai_Code_ctrm = Enum_CSV_Tana_Field.F_Tehai_Code_ICS
    F_System_Name_ctrm = Enum_CSV_Tana_Field.F_System_Name_ICS
    F_System_Spec_ctrm = Enum_CSV_Tana_Field.F_System_Spec_ICS
    F_System_Tana_No_ctrm = Enum_CSV_Tana_Field.F_System_Tana_NO_ICS
    F_Location_Text_ctrm = Enum_CSV_Tana_Field.F_Location_Text_ICS
    F_Store_Code_ctrm = Enum_CSV_Tana_Field.F_Store_Code_ICS
End Enum
'ラベル出力用一時テーブル名
Public Const T_INV_LABEL_TEMP As String = "T_INV_LABEL_TEMP"                                        'ラベル出力用の差し込み印刷用テーブルの名前
'ラベル出力用一時テーブル専用フィールド定義
Public Const F_INV_LABEL_TEMP_TEHAICODE_LENGTH As String = "F_INV_Tehaicode_Length"                 'ラベル出力のみに使用する計算列、手配コードの文字列数を格納
Public Const F_INV_LABEL_TEMP_ORDERNUM As String = "F_INV_OrderNumber"                              'ラベル出力のみに使用するオーダーNo列
Public Const F_INV_LABEL_TEMP_SAVEPOINT As String = "F_INV_Label_Savepoint"                         'ラベル出力のみに使用するSavepoint、出力リストの判別に使用する
Public Const F_INV_LABEL_TEMP_SAVE_FRENDLYNAME As String = "識別名"                                 'SavePoint出力時の Savepointフレンドリーネーム
Public Const F_INV_LABEL_TEMP_INPUT_FRENDLYNAME As String = "入力日時"                              'SavePoint出力時、InputDateフレンドリーネーム
Public Const F_INV_LABEL_TEMP_FORMSTARTTIME As String = "F_INV_Label_FormStartTime"                 'FormStartTimeを記録
Public Const F_INV_LABEL_TEMP_FRMSTART_FRENDLYNAME As String = "フォーム開始時間"                   'FormStartTimeフレンドリーネーム
'------------------------------------------------------------------------------------------------------------------------------------------------------
'DB Upsert向け定数
Public Const SQL_ALIAS_T_INVDB_Parts As String = "TDBPrts"                                          'INV_M_Partsテーブル別名定義
Public Const SQL_ALIAS_T_INVDB_Tana As String = "TDBTana"                                           'INV_M_Tanaテーブル別名定義
Public Const SQL_ALIAS_T_TEMP As String = "TTmp"                                                    '一時テーブル別名定義
Public Const SQL_ALIAS_T_SH_ZAIKO As String = "TSHZaiko"                                            '在庫情報シートテーブル名別名定義
Public Const SQL_ALIAS_T_INV_CSV As String = "TCSVTana"                                             '棚卸CSVテーブルの別名定義
Public Const SQL_ALIAS_SH_CSV As String = "SHCSV"                                                   '棚卸CSVファイルそのものの別名定義
'SQLAliasEnum
Public Enum Enum_SQL_INV_Alias
    INVDB_Parts_Alias_sia = 1
    INVDB_Tana_Alias_sia = 2
    INVDB_Tmp_Alias_sia = 3
    ZaikoSH_Alias_sia = 4
    TanaCSV_Alias_sia = 5
    SHCSV_Alias_sia = 6
End Enum
Public Const SQL_AFTER_IN_ACCDB_0FullPath As String = "[MS ACCESS;DATABASE={0};]"                   'Select From の IN""句の後に来る文字列accdb
Public Const SQL_AFTER_IN_XLSM_0FullPath As String = "[Excel 12.0 Macro;DATABASE={0};HDR=Yes;]"     'In xlsm,xlam
Public Const SQL_AFTER_IN_XLSB_0FullPath As String = "[Excel 12.0;DATABASE={0};HDR=Yes;]"           'IN xlsb
Public Const SQL_AFTER_IN_XLSX_0FullPath As String = "[Excel 12.0 xml;DATABASE={0};HDR=Yes;]"       'IN xlsx
Public Const SQL_AFTER_IN_XLS_0FullPath As String = "[Excel 8.0;DATABASE={0};HDR=Yes;]"             'IN xls
'SQL定義
'------------------------------------------------------------------------------------------
'手配コード先頭n文字リスト取得
Public Const SQL_INV_TEHAICODE_n_0TableName_1DigitNum As String = "SELECT DISTINCT LEFT(" & F_SH_ZAIKO_TEHAI_TEXT & ",{1}) FROM {0}"             '0にテーブル名を入れる
'------------------------------------------------------------------------------------------
'SH_ZaikoをT_INV_Tmpに入れる
'在庫情報シートのみ外部ファイル参照なので、IN句で指定してやる
'INの後はダブルクォーテーションふたつ、ファイル名に空白があってもエスケープする必要なし?
'T_INV_Tmpが存在していたらエラーになるので事前に削除が必要
Public Const SQL_INV_SH_TO_DB_TEMPTABLE_0Table_1INword As String = "SELECT * INTO " & T_INV_TEMP & " " & vbCrLf & _
"FROM " & vbCrLf & _
    "(SELECT * FROM {0} " & vbCrLf & _
    "IN """"{1} ) "
'
''------------------------------------------------------------------------------------------
'DISTINCT ロケーション した結果を一時テーブルに入れる
''どちらも同じDBファイル上にあるので特段IN句の指定の必要なし
'SELECT DISTINCT の結果を一時テーブルに入れる
'事前にT_INV_SELECT_TEMPの削除が必要
Public Const SQL_INV_SELECT_DISTINCT_TO_TEMPTABLE_0FieldName As String = "SELECT * INTO " & T_INV_SELECT_TEMP & " " & vbCrLf & _
"FROM ( " & vbCrLf & _
    "SELECT DISTINCT TRIM({0}) AS {0} FROM " & T_INV_TEMP & " " & vbCrLf & _
    "WHERE IIF(ISNULL(TRIM({0})),""NULL_DATA"",TRIM({0})) <> ""NULL_DATA"" AND IIF(ISNULL(TRIM({0})),""NULL_DATA"",TRIM({0})) <>  """"" & vbCrLf & _
")"
''------------------------------------------------------------------------------------------
'一時テーブルのZaikoSHをINV_M_Tanaに入れる
''一時テーブルを作成した上でのUpdateは成功
'外部DBを入力元に使う場合は、IN句はFROMの後でなければ動かないようなので、サブクエリで(SELECE * FROM Tname IN・・・) as TAlias としてやらないとダメ
Public Const SQL_INV_TEMP_TO_M_TANA_0INVDBFullPath_1LocalTimeMillisec As String = "UPDATE " & T_INV_M_Tana & " AS " & SQL_ALIAS_T_INVDB_Tana & " " & vbCrLf & _
"RIGHT JOIN ( " & vbCrLf & _
"SELECT *  FROM " & T_INV_SELECT_TEMP & " " & vbCrLf & _
"IN """"[MS ACCESS;DATABASE={0};] ) AS " & SQL_ALIAS_T_TEMP & " " & vbCrLf & _
"ON " & SQL_ALIAS_T_INVDB_Tana & "." & F_INV_TANA_SYSTEM_TEXT & " = " & SQL_ALIAS_T_TEMP & "." & F_SH_ZAIKO_TANA_TEXT & " " & vbCrLf & _
"Set " & SQL_ALIAS_T_INVDB_Tana & "." & F_INV_TANA_SYSTEM_TEXT & " = " & SQL_ALIAS_T_TEMP & "." & F_SH_ZAIKO_TANA_TEXT & "," & vbCrLf & _
SQL_ALIAS_T_INVDB_Tana & "." & PublicConst.INPUT_DATE & " = {1} " & vbCrLf & _
"WHERE " & F_INV_TANA_SYSTEM_TEXT & " Is Null"
''------------------------------------------------------------------------------------------
'T_INV_M_PartsをUpsertするSQL 入力元は T_INV_Tana と T_INV_Tmp
'3個のテーブルでJoinしてSelect,動くやつ
'SELECT TDBTana.*,TTmp.*,TDBTana.*
'FROM T_INV_M_Parts AS TDBPrts
'RIGHT JOIN (
'    T_INV_M_Tana As TDBTana
'        RIGHT JOIN (
'        SELECT * FROM  T_INV_Temp
'        IN ""[MS ACCESS;DATABASE=C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\bin\Inventory_DB\DB_Temp_Local.accdb;] ) AS TTmp
'        ON TDBTana.F_INV_Tana_System_Text = TTmp.[ロケーション])
'    ON TDBPrts.F_INV_Tehai_Code = TTmp.手配コード;
'置換サンプル
'SELECT {3}.*,{5}.*,{3}.*
'FROM {0} AS {1}
'RIGHT JOIN (
'    {2} As {3}
'        RIGHT JOIN (
'        SELECT * FROM  {4}
'        IN ""[MS ACCESS;DATABASE={6};] ) AS {5}
'        ON {3}.{7} = {5}.{8})
'    ON {1}.{9} = {5}.{10};
'
'T_INV_M_Parts                      {0}
'TDBPrts                            {1}
'T_INV_M_Tana                       {2}
'TDBTana                            {3}
'T_INV_Temp                         {4}
'TTmp                               {5}
'(CreateAfterINWord(DB_Temp.accdb)  {6}
'(ON Condition Tana and TTmp)       {7}
'(ON Condition Parts and TTmp)      {8}
'F_INV_Tana_ID                      {9}
'(SET Condition Parts and TTmp)     {10}
'(WHERE condition Parts ad TTmp)    {11}
'F_INV_Tehai_Code                   {12}
'{13} InputDate
'{14} (GetLocaltimeWithMilliSec)
Public Const SQL_INV_UPSERT_PARSTABL_FROM_TTMP_AND_TANA As String = "UPDATE  {0} AS {1} " & vbCrLf & _
"RIGHT JOIN ( " & vbCrLf & _
"   {2} As {3} " & vbCrLf & _
"   RIGHT JOIN ( " & vbCrLf & _
"       SELECT * FROM  {4} " & vbCrLf & _
"       IN """"{6} ) AS {5}" & vbCrLf & _
"   ON {7} ) " & vbCrLf & _
"ON {8} " & vbCrLf & _
"SET {1}.{9} = IIF(ISNULL({3}.{9}),-1,{3}.{9}),{1}.{13} = ""{14}"",{10} " & vbCrLf & _
"WHERE ISNULL({1}.{12}) OR {11} ;"
'T_INV_M_TanaとT_INV_M_Parts結合した汎用SELECT SQL、外部DBファイル参照は無いものとする
'{0}    (SELECT Field)
'{1}    T_INV_M_Parts
'{2}    TDBPrts
'{3}    T_INV_M_Tana
'{4}    TDBTana
'{5}    F_INV_Tana_ID
'{6}    (WHERE condition)無しの場合は空文字 "" でOK
Public Const SQL_INV_JOIN_TANA_PARTS As String = "SELECT {0} " & vbCrLf & _
"FROM {1} AS {2} " & vbCrLf & _
"   LEFT JOIN {3} AS {4} " & vbCrLf & _
"   ON {2}.{5} = {4}.{5} " & vbCrLf & _
"WHERE 1=1 {6} ;"
'Tanaマスターで、Local_textが空欄の物を一括でSystem_Textのものに設定する
'{0}    T_INV_M_Tana
'{1}    TDBTana
'{2}    (SET condition) TDBTana.F_INV_LOCAL_TEXT = TDBTana.F_INV_SYSTEM_Text
'{3}    (WHERE condition) AND TDBTana.LOCAL_TExt IS NULL
Public Const SQL_INV_TANA_SET_LOCAL_TEXT_BY_SYSTEM As String = "UPDATE {0} AS {1} " & vbCrLf & _
"SET {2} " & vbCrLf & _
"WHERE 1=1 {3}"
'------------------------------------------------------------------------------------------------
'TanaCSVをTTmpに入れて、なおかつ棚卸締切日フィールド追加して、データ入れた後に更新かけるSQLひな形
'ポイントは、JOIN の ON 条件で ANDで2個指定と、Whereで 左テーブルの棚卸締切日に Is Nullを付けること
'UPDATE T_INV_CSV AS TCSVTana
'   RIGHT JOIN (
'      SELECT * FROM T_INV_Temp
'         IN ""[MS ACCESS;DATABASE=c:\users\....] ) AS Ttmp
'   ON TCSVTana.手配コード = Ttmp.手配コード AND TCSVTana.棚卸締切日 = Ttmp.棚卸締切日
'SET TCSVTana.棚卸締切日 = "2022/02/02",TCSVTana.----
'WHERE TCSVTana.棚卸締切日 Is Null
'{0}    T_INV_CSV
'{1}    TCSVTana
'{2}    T_INV_Temp
'{3}    (After IN Word)
'{4]    Ttmp
'{5}    手配コード
'{6]    (SET condition)
'{7}    棚卸締切日
Public Const SQL_INV_TMP_TO_CSVTANA As String = "UPDATE {0} AS {1} " & vbCrLf & _
"   RIGHT JOIN (" & vbCrLf & _
"       SELECT * FROM {2} " & vbCrLf & _
"           IN """"{3} ) AS {4} " & vbCrLf & _
"   ON {1}.{5} = {4}.{5} AND {1}.{7} = {4}.{7} " & vbCrLf & _
"SET {6} " & vbCrLf & _
"WHERE {1}.{7} Is Null"
'------------------------------------------------------------------------------------------------
'DBからCSV(xls)ファイルにデータセットするSQL
'カレントディレクトリはシートファイルのものにする
'UPDATE ['SIZ_TANAOROSI_HYO - コピー_Local$'_xlnm#_FilterDatabase] AS SHCSV
'    RIGHT JOIN (
'        SELECT * FROM T_INV_CSV
'        IN "" [MS ACCESS;DATABASE=C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\bin\Inventory_DB\INV_Manege_Local.accdb]
'        WHERE 棚卸締切日 = "2022/01/12"
'        ) AS TCSVTana
'    ON SHCSV.手配コード = TCSVTana.手配コード
'Set SHCSV.現品残 = TCSVTana.現品残
'{0}    (Sheet Table Name)
'{1}    SHCSV
'{2}    T_INV_CSV
'{3}    (After In Word Default DB)
'{4}    棚卸締切日
'{5}    (2022/01/12 EndDay)
'{6}    TCSVTana
'{7}    手配コード
'{8}    現品残
Public Const SQL_INV_DB_TO_CSV As String = "UPDATE {0} AS {1}" & vbCrLf & _
"    RIGHT JOIN (" & vbCrLf & _
"        SELECT * FROM {2}" & vbCrLf & _
"        IN """" {3}" & vbCrLf & _
"        WHERE {4} = ""{5}""" & vbCrLf & _
"        ) AS {6}" & vbCrLf & _
"    ON {1}.{7} = {6}.{7}" & vbCrLf & _
"Set {1}.{8} = {6}.{8}"
'------------------------------------------------------------------------------------------------
'ラベル出力用一時テーブル作成SQL
'CREATE TABLE T_INV_LABEL_TEMP (
'    F_INV_Tana_Local_Text CHAR(10),F_INV_Tehai_Code CHAR(50),
'    F_INV_Label_Name_1 CHAR(18),F_INV_Label_Name_2 CHAR(18),F_INV_Label_Remark_1 CHAR(18),F_INV_Label_Remark_2 CHAR(18),InputDate CHAR(23)
')
'{0}    T_INV_LABEL_TEMP
'{1}    F_INV_Tana_Local_Text
'{2}    F_INV_Tehai_Code
'{3}    F_INV_Label_Name_1
'{4}    F_INV_Label_Name_2
'{5}    F_INV_Label_Remark_1
'{6}    F_INV_Label_Remark_2
'{7}    InputDate
'{8}    INV_CONST.F_INV_LABEL_TEMP_TEHAICODE_LENGTH
'{9}    INV_CONST.F_INV_LABEL_TEMP_ORDERNUM
'{10}   INV_CONST.F_INV_LABEL_TEMP_SAVEPOINT
'{11}   INV_CONST.F_INV_LABEL_TEMP_FORMSTARTTIME
Public Const SQL_INV_CREATE_LABEL_TEMP_TABLE As String = "CREATE TABLE {0} (" & vbCrLf & _
"    {10} CHAR(23),{11} CHAR(23),{1} CHAR(15),{2} CHAR(50),{8} LONG," & vbCrLf & _
"    {3} CHAR(18),{4} CHAR(18),{5} CHAR(18),{6} CHAR(18),{9} CHAR(9),{7} CHAR(23)" & vbCrLf & _
")"
'------------------------------------------------------------------------------------------------
'Label_Temp SavePoint一覧出力SQL
'{0}    F_inv_Label_Savepoint
'{1}    FormStartTime
'{2}    T_inv_Label_Temp
'{3}    INV_CONST.F_INV_LABEL_TEMP_SAVE_FRENDLYNAME
'{4}    INV_CONST.F_INV_LABEL_TEMP_FRMSTART_FRENDLYNAME
Public Const SQL_SELECT_SAVEPOINT As String = "SELECT {0} AS {3},{1} AS {4}  FROM {2} " & vbCrLf & _
"GROUP BY {0},{1} " & vbCrLf & _
"ORDER BY {0} DESC,{1} DESC"
'------------------------------------------------------------------------------------------------
'BinLabel MailMerge用基礎データ取得
'{0}    INV_CONST.T_INV_LABEL_TEMP
'{1}    (MailMerge Where)
'{2}    INV_CONST.T_INV_SELECT_TEMP
Public Const SQL_LABEL_MAILMERGE_DEFAULT As String = "SELECT * INTO {2} FROM [{0}] " & vbCrLf & _
"WHERE {1}"
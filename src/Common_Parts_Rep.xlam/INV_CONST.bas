Attribute VB_Name = "INV_CONST"
'''在庫管理関係の定数を定義する
Option Explicit
'在庫情報DBに関する定数
Public Const INV_DB_FILENAME As String = "INV_Manege.accdb"                     '在庫情報のDBファイル名
Public Const T_INV_TEMP As String = "T_INV_Temp"                                'INVDBの一時テーブル名
Public Const T_INV_SELECT_TEMP As String = "T_INV_Select_Temp"                  'Selectした結果を格納するテーブル名、一旦テーブルに格納しないと
                                                                                '更新可能なクエリが・・・とか言われるため
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
    F_INV_TANA_ID_IMT = 1
    F_INV_Tana_Local_Text_IMT = 2
    F_INV_Tana_System_Text_IMT = 3
    F_INV_TIET_Delivary_IMT = 4
    F_InputDate_IMT = 5
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
End Enum
'------------------------------------------------------------------------------------------------------------------------------------------------------
'SQL定義
'手配コード先頭n文字リスト取得
Public Const SQL_INV_TEHAICODE_n_0TableName_1DigitNum As String = "SELECT DISTINCT LEFT(" & F_SH_ZAIKO_TEHAI_TEXT & ",{1}) FROM {0}"             '0にテーブル名を入れる
'DB Upsert向け定数
Public Const SQL_ALIAS_T_INVDB_Parts As String = "TDBPrts"                                          'INV_M_Partsテーブル別名定義
Public Const SQL_ALIAS_T_INVDB_Tana As String = "TDBTana"                                           'INV_M_Tanaテーブル別名定義
Public Const SQL_ALIAS_T_TEMP As String = "TTmp"                                                    '一時テーブル別名定義
Public Const SQL_ALIAS_T_SH_ZAIKO As String = "TSHZaiko"                                            '在庫情報シートテーブル名別名定義
'SQLAliasEnum
Public Enum Enum_SQL_INV_Alias
    INVDB_Parts_Alias_sia = 1
    INVDB_Tana_Alias_sia = 2
    INVDB_Tmp_Alias_sia = 3
    ZaikoSH_Alias_sia = 4
End Enum
Public Const SQL_AFTER_IN_ACCDB_0FullPath As String = "[MS ACCESS;DATABASE={0};]"                   'Select From の IN""句の後に来る文字列accdb
Public Const SQL_AFTER_IN_XLSM_0FullPath As String = "[Excel 12.0 Macro;DATABASE={0};HDR=Yes;]"     'In xlsm,xlam
Public Const SQL_AFTER_IN_XLSB_0FullPath As String = "[Excel 12.0;DATABASE={0};HDR=Yes;]"           'IN xlsb
Public Const SQL_AFTER_IN_XLSX_0FullPath As String = "[Excel 12.0 xml;DATABASE={0};HDR=Yes;]"       'IN xlsx
Public Const SQL_AFTER_IN_XLS_0FullPath As String = "[Excel 8.0;DATABASE={0};HDR=Yes;]"             'IN xls
'在庫情報シートのみ外部ファイル参照なので、IN句で指定してやる
'INの後はダブルクォーテーションふたつ、ファイル名に空白があってもエスケープする必要なし?
'SELECT TSHZaiko.手配コード,TDBTana.F_INV_Tana_ID,TDBTana.F_INV_Tana_Local_Text,TDBTana.F_INV_Tana_System_Text
'FROM    (
'    SELECT * FROM T_INV_M_Tana
'    IN ""[MS ACCESS;DATABASE=R:\Tmp\Patacchi\Test Dir\INV_Manege_Local.accdb;]
'    ) AS TDBTana
'RIGHT JOIN (
'    SELECT * FROM [在庫情報$FilterDatabase]
'    IN ""[Excel 12.0;DATABASE=R:\Tmp\Patacchi\Test Dir\Zaiko_0_Local.xls;]
'    ) AS TSHZaiko
'ON TDBTana.F_INV_Tana_System_Text = TSHZaiko.ロケーション
'WHERE NOT TDBTana.F_INV_Tana_ID IS NULL;
''------------------------------------------------------------------------------------------
''外部データを新規テーブルとしてインポートする
''T_Tempが存在していたらエラーになるので事前に削除が必要
Public Const SQL_INV_SH_TO_DB_TEMPTABLE_0Table_1INword As String = "SELECT * INTO " & T_INV_TEMP & " " & vbCrLf & _
"FROM " & vbCrLf & _
    "(SELECT * FROM {0} " & vbCrLf & _
    "IN """"{1} ) "
'
''------------------------------------------------------------------------------------------
''どちらも同じDBファイル上にあるので特段IN句の指定の必要なし
'SELECT * INTO T_INV_Temp_Select
'FROM (
'    SELECT DISTINCT ロケーション FROM T_INV_Temp
')
'SELECT DISTINCT の結果を一時テーブルに入れる
'事前にT_INV_SELECT_TEMPの削除が必要
Public Const SQL_INV_SELECT_DISTINCT_TO_TEMPTABLE_0FieldName As String = "SELECT * INTO " & T_INV_SELECT_TEMP & " " & vbCrLf & _
"FROM ( " & vbCrLf & _
    "SELECT DISTINCT {0} FROM " & T_INV_TEMP & vbCrLf & _
")"
''------------------------------------------------------------------------------------------
''一時テーブルを作成した上でのUpdateは成功
'UPDATE T_INV_M_Tana AS TDBTana
'RIGHT JOIN (
'SELECT * FROM T_INV_Temp
'IN ""[MS ACCESS;DATABASE=C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\bin\Inventory_DB\DB_Temp_Local.accdb;] ) AS TDBTemp
'ON TDBTana.F_INV_Tana_System_Text = TDBTemp.ロケーション
'Set TDBTana.F_INV_Tana_System_Text = TDBTemp.ロケーション,
'TDBTana.InputDate = "2022-01-25T16.20:00.010"
'WHERE F_INV_Tana_System_Text Is Null
Public Const SQL_INV_TEMP_TO_M_TANA_0INVDBFullPath_1LocalTimeMillisec As String = "UPDATE " & T_INV_M_Tana & " AS " & SQL_ALIAS_T_INVDB_Tana & " " & vbCrLf & _
"RIGHT JOIN ( " & vbCrLf & _
"SELECT *  FROM " & T_INV_SELECT_TEMP & " " & vbCrLf & _
"IN """"[MS ACCESS;DATABASE={0};] ) AS " & SQL_ALIAS_T_TEMP & " " & vbCrLf & _
"ON " & SQL_ALIAS_T_INVDB_Tana & "." & F_INV_TANA_SYSTEM_TEXT & " = " & SQL_ALIAS_T_TEMP & "." & F_SH_ZAIKO_TANA_TEXT & " " & vbCrLf & _
"Set " & SQL_ALIAS_T_INVDB_Tana & "." & F_INV_TANA_SYSTEM_TEXT & " = " & SQL_ALIAS_T_TEMP & "." & F_SH_ZAIKO_TANA_TEXT & "," & vbCrLf & _
SQL_ALIAS_T_INVDB_Tana & "." & PublicConst.INPUT_DATE & " = {1} " & vbCrLf & _
"WHERE " & F_INV_TANA_SYSTEM_TEXT & " Is Null"
''------------------------------------------------------------------------------------------
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
'
'T_INV_M_Parts   {0}
'TDBPrts {1}
'T_INV_M_Tana {2}
'TDBTana     {3}
'T_INV_Temp  {4}
'TTmp    {5}
'C:\Users\q3005sbe\AppData\Local\Rep\InventoryManege\bin\Inventory_DB\DB_Temp_Local.accdb    {6}
'F_INV_Tana_System_Text  {7}
'ロケーション    {8}
'F_INV_Tehai_Code    {9}
'手配コード  {10}
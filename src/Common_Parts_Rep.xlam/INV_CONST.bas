Attribute VB_Name = "INV_CONST"
'''在庫管理関係の定数を定義する
Option Explicit
'在庫情報DBに関する定数
Public Const INV_DB_FILENAME As String = "INV_Manege.accdb"                     '在庫情報のDBファイル名
'部品（手配コード）マスターテーブルの定数
Public Const T_INV_PARTS_MASTER As String = "T_INV_M_Parts"                     '手配コードマスターのテーブル名
Public Const F_INV_TEHAI_ID As String = "F_INV_Tehai_ID"                        '手配コードのID、各テーブルにはこの値を設定する
Public Const F_INV_TEHAI_TEXT As String = "F_INV_Tehai_Code"                    '手配コード
Public Const F_INV_MANEGE_SECTON As String = "F_INV_Manege_Section"             '管理課
Public Const F_INV_SYSTEM_TANA_NO As String = "F_INV_System_Tana_No"            'システム側の棚番号
Public Const F_INV_KISHU As String = "F_INV_Kishu"                              '機種名
Public Const F_INV_STORE_CODE As String = "F_INV_Store_Code"                    '貯蔵記号
Public Const F_INV_DELIVER_LOT As String = "F_INV_Deliver_Lot"                  '払い出しロット
Public Const F_INV_FILL_LOT As String = "F_INV_Fill_Lot"                        '補充ロット
Public Const F_INV_LEAD_TIME As String = "F_INV_Lead_Time"                      'リードタイム
Public Const F_INV_ORDER_AMOUNT As String = "F_INV_Order_Amount"                '発注数
Public Const F_INV_ORDER_REMAIN As String = "F_INV_Order_Remain"                '発注残数
Public Const F_INV_STOCK_AMOUNT As String = "F_INV_Stock_Amount"                '在庫数
Public Const F_INV_TANA_ID As String = "F_INV_Tana_ID"                          '棚番号のID、実態は棚番マスターから引っ張ること
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
'棚番のみは共有できないので、独自に数字を降る
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
    F_Order_Remail_ShZ = Enum_INV_M_Parts.F_Order_Remain_IMPrt
    F_Stock_Amount_ShZ = Enum_INV_M_Parts.F_Stock_Amount_IMPrt
    F_System_Name_ShZ = Enum_INV_M_Parts.F_System_Name_IMPrt
    F_System_Spec_ShZ = Enum_INV_M_Parts.F_System_Spec_IMPrt
    F_Store_Unit_ShZ = Enum_INV_M_Parts.F_Store_Unit_IMPrt
    F_System_Description_ShZ = Enum_INV_M_Parts.F_System_Description_IMPrt
    F_Manege_Section_Sub_ShZ = Enum_INV_M_Parts.F_Manege_Section_Sub_IMPrt
    '棚番テキストのみこちらで独自に設定する
    F_Tana_Text_ShZ = 101
End Enum
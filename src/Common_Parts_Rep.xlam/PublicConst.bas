Attribute VB_Name = "PublicConst"
Option Explicit
'データベース全般の定数定義
'DBファイルの拡張子
Public Const DB_FILE_EXETENSION1 As String = "accdb"            'DBファイルの拡張子その１(lcaseにして）
'DBファイルの列挙体
Public Enum DB_file_exetension
    accdb_dext = 0
End Enum
''-------リスト表示のための定数定義
''MS ゴシック（等幅）文字サイズ9ptの場合
'Public Const sglChrLengthToPoint = 4.1
'Public Const longMINIMULPOINT = 50
''フィールド追加用SQL定型文
'Public Const strOLDAddField1_NextTableName     As String = "ALTER TABLE """        '追加の最初、この次にテーブル名が入る
'Public Const strOLDAddField2_NextFieldName     As String = """ ADD COLUMN """      '二番目、この次にフィールド名が入る
'Public Const strOLDAddField3_Text_Last         As String = """ TEXT;"              '最後、ただしTEXT型の場合
'Public Const strOLDAddField3_Numeric_Last      As String = """ NUMERIC;"           '数値の場合の最後
'Public Const strOLDAddField3_JSON_Last         As String = """ JSON;"              'JSONラスト
''インデックス追加用SQL定型文
'Public Const strOLDIndex1_NextTable            As String = "CREATE INDEX IF NOT EXISTS ""ixJob_"
'Public Const strOLDIndex2_NextTable            As String = """ ON """
'Public Const strOLDIndex3_Field1               As String = """ ("""
'Public Const strOLDIndex4_FieldNext            As String = """ ASC ,"""            '複数フィールドに対して実行する場合は、以後これの繰り返し
'Public Const strOLDIndex5_Last                 As String = """ ASC);"
''テーブル追加用SQL定型文
'Public Const strTable1_NextTable                As String = "CREATE TABLE IF NOT EXISTS " 'CRLF付加、およびフィールド制約追加対応テンプレ
'Public Const strTable2_Next1stField             As String = " (" & vbCrLf           'CRLF対応作成テンプレ、こちらを使う場合はAddQuoteを使ってエスケープ処理すること
''フィールド定義、フィールド名（クオート）→3(型名)→[Append](各種制約、あれば)→(EndRow)→（次のフィールドがあれば）フィールド名（クオート）→型名・・・の流れ
''1 テーブル名 2 最初のフィールド 3（型名) （次があるなら）4 フィールド名・・・　（最後なら）5
'Public Const strTable3_TEXT                     As String = " TEXT "                '前がTEXT
'Public Const strTable3_NUMERIC                  As String = " NUMERIC "             '前がNUMERIC
'Public Const strTable3_JSON                     As String = " JSON "                '前がJSON
'Public Const strTable_NotNull                   As String = " NOT NULL "            'NOT NULL制約追加
'Public Const strTable_Unique                    As String = " UNIQUE "              'UNIQUE制約追加
'Public Const strTable_Default                   As String = " DEFAULT "             'DEFAULT追加、この後にデフォルト値をクオート処理して追加すること
'Public Const strTable4_EndRow                   As String = "," & vbCrLf            '行の終わり、まだ続きがある場合
'Public Const strTable4_5_PrimaryKey             As String = "PRIMARY KEY("          'PrimaryKeyの指定をこの後に続ける
'Public Const strTable4_6_EndPrimary             As String = ")" & vbCrLf            'PrimaryKey等のカッコ閉じ
'Public Const strTable5_EndSQL                   As String = ");" & vbCrLf           'SQL文の終わり
'Public Const strOLDAddTable1_NextTable          As String = "CREATE TABLE IF NOT EXISTS """ 'テーブル追加用定型文ここから
'Public Const strOLDAddTable2_Field1_Next_Field  As String = """ ("""                'フィールドの最初だけこいつを使う、次に最初のフィールド名
'Public Const strOLDAddTable_TEXT_Next_Field     As String = """ TEXT,"""            '紛らわしいけど、「前」がText型の場合こっちを使う、次にフィールド名が続く
'Public Const strOLDAddTable_TEXT_UNIQUE_Next_Field As String = """ TEXT UNIQUE,"""  '前がTEXT かつ UNIQUEの場合
'Public Const strOLDAddTable_NUMELIC_Next_Field  As String = """ NUMERIC,"""         '「前」がNumericの場合はこっち
'Public Const strOLDAddTable_Text_Last           As String = """ TEXT);"             'メンドウなので、最後はTextで終わらせて・・・
'Public Const strOLDAddTable_Numeric_Last        As String = """ NUMERIC);"          '一応数値型で終わるやつも
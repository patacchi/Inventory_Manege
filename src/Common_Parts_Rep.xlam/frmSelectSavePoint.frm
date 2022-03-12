VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectSavePoint 
   Caption         =   "印刷リスト選択画面"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6210
   OleObjectBlob   =   "frmSelectSavePoint.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectSavePoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Label_TempのSavepointを選択するフォーム
Option Explicit
Private StopEvents As Boolean
Private Enum Enum_SortField
    SavePoint = 0
    InputDate = 1
End Enum
'ListのInitializeは呼び出し元で行う
'''キャンセルボタン
Private Sub btnCancel_Click()
    '''frmBinLabelのvarstrarrSelectedSavePointにEmptyをセットして、自身をUnload
    frmBinLabel.varstrarrSelectedSavepoint = Empty
    Unload Me
    Exit Sub
End Sub
Private Sub btnCompSelect_Click()
    '選択されたそうで
    Dim strarrSelectedSavepoint() As String 'SavePoint格納用配列
    Dim strarrInputDate() As String         'InputDate格納用配列
    Dim longListRow As Long                 'Listのカレント行を格納
    Dim longSelectedRows As Long            '選択された項目の数を格納
    longSelectedRows = 0
    'lstBoxSavePoint.Listの全項目をループ
    For longListRow = LBound(lstBoxSavePoint.List) To UBound(lstBoxSavePoint.List)
        Select Case True
        Case longListRow = LBound(lstBoxSavePoint.List)
            '最初の行の場合
            '最初の行はタイトル行なので何もしない
        Case lstBoxSavePoint.Selected(longListRow)
            '選択されている項目の場合
            'longSelectedRowsをインクリメント
            longSelectedRows = longSelectedRows + 1
            '結果配列Redim Preserve
            ReDim Preserve strarrSelectedSavepoint(longSelectedRows - 1)
            ReDim Preserve strarrInputDate(longSelectedRows - 1)
            'Listの項目をセット
            'SavePoint
            strarrSelectedSavepoint(longSelectedRows - 1) = CStr(lstBoxSavePoint.List(longListRow, 0))
            'InputDate
            strarrInputDate(longSelectedRows - 1) = CStr(lstBoxSavePoint.List(longListRow, 1))
        End Select
    Next longListRow
    If longSelectedRows < 1 Then
        MsgBox "選択された項目が有りませんでした。最低でも1項目選択して下さい"
        Exit Sub
    End If
    'frmBinLabelに取得した結果をセットし、自身をUnload
    Dim longArrayRowCounter As Long
    '結果格納用配列を定義
    Dim arrstrResult() As String
    ReDim arrstrResult(UBound(strarrSelectedSavepoint), 1)
    '結果のSavePoint要素分ループし、SavePointとInputDateをセット
    For longArrayRowCounter = LBound(strarrSelectedSavepoint) To UBound(strarrSelectedSavepoint)
        'SavePointセット
        arrstrResult(longArrayRowCounter, 0) = strarrSelectedSavepoint(longArrayRowCounter)
        'InputDateセット
        arrstrResult(longArrayRowCounter, 1) = strarrInputDate(longArrayRowCounter)
    Next longArrayRowCounter
    'frmBinLabelに結果の配列をセット
    frmBinLabel.varstrarrSelectedSavepoint = arrstrResult
    Unload Me
    Exit Sub
End Sub
'InputDate順に並び替え
Private Sub btnOrderByInputDate_Click()
    SortList Enum_SortField.InputDate
End Sub
'SavePoint順に並び替え
Private Sub btnOrderBySavePoint_Click()
    SortList Enum_SortField.SavePoint
End Sub
'''lstbox_Change
'''先頭行だったら全選択、全解除の動きをする
Private Sub lstBoxSavePoint_Change()
    If StopEvents Then
        'イベント停止フラグ立ってたら抜ける
        Exit Sub
    End If
    If lstBoxSavePoint.ListCount < 2 Then
        'リストが2行未満(タイトルだけ?)の場合は抜ける
        MsgBox "有効なデータがありませんでした"
        Exit Sub
    End If
    '最初の行が選択されたら、他をTrueにするかFalseにするか決める
    Select Case True
    Case lstBoxSavePoint.ListIndex = 0
        '最初の行が選択された
        'リスト全ての行をループ、ただし最初の行以外
        '全ての行のSelectedを最初の行と同じにする
        'イベント停止する
        StopEvents = True
        Dim longArrayCount As Long
        For longArrayCount = (LBound(lstBoxSavePoint.List) + 1) To UBound(lstBoxSavePoint.List)
            lstBoxSavePoint.Selected(longArrayCount) = lstBoxSavePoint.Selected(0)
        Next longArrayCount
        'イベント再開する
        StopEvents = False
    End Select
End Sub
Private Sub SortList(argEnumField As Enum_SortField)
    'リストを並び変える
    '現在の情報をレコードセットに格納
    Dim rsSavePoint As ADODB.Recordset
    Set rsSavePoint = New ADODB.Recordset
    'フィールド名定義は1行目の物をフィールド名、型はStringとする
    'SavePointフィールド追加
    rsSavePoint.Fields.Append Name:=CStr(lstBoxSavePoint.List(0, 0)), Type:=adWChar, DefinedSize:=23
    'InptDateフィールド追加
    rsSavePoint.Fields.Append Name:=CStr(lstBoxSavePoint.List(0, 1)), Type:=adWChar, DefinedSize:=23
    'rsをOpenする
    If rsSavePoint.State = ObjectStateEnum.adStateClosed Then
        rsSavePoint.Open
    End If
    'データを追加する、実際のデータはListの2行目から
    Dim longListRowCount As Long
    For longListRowCount = 1 To UBound(lstBoxSavePoint.List)
        rsSavePoint.AddNew
        'SavePoint
        rsSavePoint.Fields(CStr(lstBoxSavePoint.List(0, 0))).Value = lstBoxSavePoint.List(longListRowCount, 0)
        'InputDate
        rsSavePoint.Fields(CStr(lstBoxSavePoint.List(0, 1))).Value = lstBoxSavePoint.List(longListRowCount, 1)
        'このループのRS確定
        rsSavePoint.Update
    Next longListRowCount
    'Sort実行
    rsSavePoint.Sort = CStr(lstBoxSavePoint.List(0, CLng(argEnumField))) & " DESC"
    '結果格納用配列定義
    Dim varArr As Variant
    ReDim varArr(UBound(lstBoxSavePoint.List), 1)
    '1行目はタイトル行固定
    varArr(0, 0) = CStr(lstBoxSavePoint.List(0, 0))
    varArr(0, 1) = CStr(lstBoxSavePoint.List(0, 1))
    '行カウンター初期化、2行目から
    longListRowCount = 1
    'rsMoveFirst
    'rsをループして、EOFになるまでデータ追加
    rsSavePoint.MoveFirst
    Do
        'SavePoint
        varArr(longListRowCount, 0) = rsSavePoint.Fields(0).Value
        'InputDate
        varArr(longListRowCount, 1) = rsSavePoint.Fields(1).Value
        'longListRowCountインクリメント
        longListRowCount = longListRowCount + 1
        'MoveNext
        rsSavePoint.MoveNext
    Loop While Not rsSavePoint.EOF
    '結果をListBoxに適用
    lstBoxSavePoint.Clear
    lstBoxSavePoint.List = varArr
    'rsを解放
    Set rsSavePoint = Nothing
    Exit Sub
End Sub
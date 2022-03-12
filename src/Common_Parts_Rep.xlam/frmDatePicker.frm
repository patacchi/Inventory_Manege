VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   Caption         =   "日付選択フォーム"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   OleObjectBlob   =   "frmDatePicker.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''イベント共通化関係変数定義
'''Dateの初期値は cdbl(date)で 0になる
'共通イベントクラス格納コレクション
Private colcommonEvents As Collection
'''メンバ変数定義
Private dateCurrentMonth As Date                            'フォームが現在扱っている年月日(日は1を指定したもの)を格納する変数
Private StopEvents As Boolean                               'イベント制御フラグ
'共通イベントハンドラで結果を書き込むための変数
Public vardateDayArray_CurrentYearMonth As Variant              '現在の年月を元に取得したラベル個数分のDate配列
'''-------------------------------------------------------------------------------------------------------------
'''イベント
'Userform
'Initialize
Private Sub UserForm_Initialize()
    ConstRuctor
End Sub
'''前の月へ Click
Private Sub btnPreviousMonth_Click()
    '現在の年月を取得
    Dim dateCurent As Date
    dateCurent = DateSerial(CInt(cmbBox_Year.Text), CInt(cmdBox_Month.Text), 1)
    '年月にMonth-1したものを設定
    cmbBox_Year.Text = Year(DateSerial(Year(dateCurent), Month(dateCurent) - 1, 1))
    cmdBox_Month.Text = Month(DateSerial(Year(dateCurent), Month(dateCurent) - 1, 1))
End Sub
'''次の月へ Click
Private Sub btnNextMonth_Click()
    '現在の年月を取得
    Dim dateCurent As Date
    dateCurent = DateSerial(CInt(cmbBox_Year.Text), CInt(cmdBox_Month.Text), 1)
    '年月にMonth+1したものを設定
    cmbBox_Year.Text = Year(DateSerial(Year(dateCurent), Month(dateCurent) + 1, 1))
    cmdBox_Month.Text = Month(DateSerial(Year(dateCurent), Month(dateCurent) + 1, 1))
End Sub
'''-------------------------------------------------------------------------------------------------------------
'''メソッド
'''コンストラクタ
Private Sub ConstRuctor()
    'イベント停止する
    StopEvents = True
    '最初にFormCommonの結果Dateを初期化する
    FormCommon.datePickerResult = Empty
    '初期日時として当月1日を設定
    dateCurrentMonth = DateSerial(Year(Now()), Month(Now()), 1)
    '年月日のボックスを設定
    '年 前後2年間
    Dim intYear(4) As Integer
    Dim longArrayCounter As Long
    For longArrayCounter = 0 To UBound(intYear)
        intYear(longArrayCounter) = Year(dateCurrentMonth) - 2 + longArrayCounter
    Next longArrayCounter
    cmbBox_Year.List = intYear
    '月 1-12
    Dim intMonth(11) As Integer
    For longArrayCounter = 0 To UBound(intMonth)
        intMonth(longArrayCounter) = longArrayCounter + 1
    Next longArrayCounter
    cmdBox_Month.List = intMonth
    '当月のものにする
    cmdBox_Month.Text = Month(dateCurrentMonth)
    '日付ボックスはイベント共通化イベントハンドラで結果が確定したらそちらで設定される
    'イベント共通化ハンドルの処理
    '共通イベントハンドラ イベントクラス格納用コレクションの初期化
    If colcommonEvents Is Nothing Then
        Set colcommonEvents = New Collection
    End If
    setCommonEventsContrl
    '初期設定のため、年を当年のものにし、Changeイベントを発生させる
    'テキストを当年のものにする
    cmbBox_Year.Text = Year(dateCurrentMonth)
End Sub
''’コンボボックスに入力された年月から日付ボックスに設定する配列を得る
'''-------------------------------------------------------------------------------------------------------
'''Return   Integer() 配列
Private Function GetDayArrayFromCurrentMonth() As Integer()
    '年月に数字以外が入っていないかチェック
    If Not IsNumeric(cmbBox_Year.Text) Or Not IsNumeric(cmdBox_Month.Text) Then
        '年月いずれかのコンボボックスに入力された日付が数字として認識できない場合
        MsgBox "年月ボックスいずれかに数字以外が入力されました"
        GetDayArrayFromCurrentMonth = Empty
        GoTo CloseAndExit
    End If
    '日付として変換できるかチェック
    If Not IsDate(DateSerial(cmbBox_Year.Text, cmdBox_Month.Text, 1)) Then
        'コンボボックスの年月の数字が日付に変換できないとき
        MsgBox "入力された数値が日付として認識できませんでした"
        GetDayArrayFromCurrentMonth = Empty
        GoTo CloseAndExit
    End If
    Dim intDay() As Integer
    '翌月の(Month + 1)前日(day 0)(=当月の末日) -1 が配列のサイズとなる
    ReDim intDay(Day(DateSerial(CInt(cmbBox_Year.Text), CInt(cmdBox_Month.Text) + 1, 0)) - 1)
    Dim longArrayRowCounter As Long
    For longArrayRowCounter = 0 To UBound(intDay)
        intDay(longArrayRowCounter) = longArrayRowCounter + 1
    Next longArrayRowCounter
    '結果を返す
    GetDayArrayFromCurrentMonth = intDay
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "GetDayArrayFromCurrentMonth code: " & Err.Number & " Description: " & Err.Description
    GetDayArrayFromCurrentMonth = Empty
    GoTo CloseAndExit
CloseAndExit:
    Exit Function
End Function
'''-------------------------------------------------------------------------------------------------------
'''イベントハンドラ共通化クラスに登録する
Private Sub setCommonEventsContrl()
    Dim Ctrleach As MSForms.Control
    '正規表現オブジェクトのインスタンス取得
#If refRegExp Then
    'RegExpが参照設定されているとき
    Dim objRegExp As RegExp
    Set objRegExp = New RegExp
#Else
    '参照設定されていないとき(遅延バインディング)
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
#End If
    'lbl_Day?? 数字2桁
    objRegExp.Pattern = FormCommon.LABEL_DAY_PRIFIX & "[0-9]{2}"
    'イベント共通化クラスのインスタンスを得る
    Dim clsCommonEventsDatePicker As clsCommonEvents
    'DatePickerの全コントロールをループ
    For Each Ctrleach In frmDatePicker.Controls
        Select Case True
        Case objRegExp.Test(Ctrleach.Name), Ctrleach.Name = cmbBox_Year.Name, Ctrleach.Name = cmdBox_Month.Name
            '日付のラベルコントロールだった(名前の正規表現一致による)
            'または年月コンボボックスだった
            Set clsCommonEventsDatePicker = New clsCommonEvents
            'イベント処理クラスのWithEvents変数に現在のコントロールをセット
            Set clsCommonEventsDatePicker.commonControl = Ctrleach
            'イベント共通化クラスコレクションに追加
            colcommonEvents.Add clsCommonEventsDatePicker
            Set clsCommonEventsDatePicker = Nothing
        End Select
    Next Ctrleach
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CDbl(FormCommon.datePickerResult) = 0 Then
        '閉じるときに結果が初期値だったらそのまま閉じていいか聞く
        Dim longMsgBoxResult As Long
        longMsgBoxResult = MsgBox("日付が選択されていませんが、このまま閉じますか？", vbYesNo)
        If longMsgBoxResult = vbYes Then
            'そのまま閉じてもいい
            Cancel = 0
        Else
            '閉じちゃダメ
            Cancel = 1
        End If
    End If
End Sub
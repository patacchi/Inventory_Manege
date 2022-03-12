VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''イベントプロシージャ共通化のためのクラス
'''フォーム名、コントロール種別ごとに別の変数を用意する
Option Explicit
'''---------------------------------------------------------------------------
'frmDatePickerのLabelの共通イベントハンドラ補足用の変数
Private WithEvents common_DatePicker_Label As MSForms.Label
Attribute common_DatePicker_Label.VB_VarHelpID = -1
Private WithEvents common_DatePicker_cmbBox As MSForms.ComboBox
Attribute common_DatePicker_cmbBox.VB_VarHelpID = -1
'''---------------------------------------------------------------------------
'外部公開用(プロパティ)の汎用Control変数
Private pControl As MSForms.Control
'''---------------------------------------------------------------------------
'プロパティ
Property Get commonControl() As MSForms.Control
    Set commonControl = pControl
End Property
Property Set commonControl(argCtrl As MSForms.Control)
    On Error GoTo ErrorCatch
    'コントロールの種類により処理を分岐
    If argCtrl Is Nothing Then
        '引数がNothingだったら抜ける
        GoTo CloseAndExit
        Exit Property
    End If
    Select Case TypeName(argCtrl)
    Case "Label"
        'ラベルの場合
        '更にフォーム名で処理を分岐
        Select Case argCtrl.Parent.Name
        Case frmDatePicker.Name
            'DatePickerのラベルだった
            Set pControl = argCtrl
            'イベント補足用変数にセット(以後イベントキャッチできるようになるはず)
            Set common_DatePicker_Label = Me.commonControl
            GoTo CloseAndExit
        End Select
    Case "ComboBox"
        'コンボボックスの場合
        '更にフォーム名で処理を分岐
        Select Case argCtrl.Parent.Name
        Case frmDatePicker.Name
            'DatePickerのコンボボックスだった
            Set pControl = argCtrl
            Set common_DatePicker_cmbBox = Me.commonControl
            GoTo CloseAndExit
        End Select
    End Select
    'ここまで抜けてきたという事は未対応の組み合わせだった
    MsgBox "次のコントロールは共通化定義されていません" & vbCrLf & _
    "コントロール名: " & TypeName(argCtrl) & " " & argCtrl.Name & vbCrLf & _
    "親フォーム名： " & argCtrl.Parent.Name
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "set commonControl code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Property
End Property
'''---------------------------------------------------------------------------
'''共通イベントハンドラ記述
'''DatePicker Label Click
'''親フォームの年月ボックスより数値を取得し、自身のラベルの番号
Private Sub common_DatePicker_Label_Click()
#If refRegExp Then
    'RegExp参照済み
    Dim objRegExp As RegExp
    Set objRegExp = New RegExp
#Else
    'RegExp参照なし
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
#End If
    '正規表現パターン設定
    'プリフィックスと数字部分2グループに分ける
    objRegExp.Pattern = "(" & FormCommon.LABEL_DAY_PRIFIX & ")([0-9]{2})"
    If Not objRegExp.Test(Me.commonControl.Name) Then
        MsgBox "日付ラベル以外のラベルです。処理を中断します"
        GoTo CloseAndExit
    End If
    '数値部分を取得
#If refRegExp Then
    'regexp参照されている場合
    Dim MatchCollection As MatchCollection
#Else
    '参照無しの場合
    Dim MatchCollection As Object
#End If
    Set MatchCollection = objRegExp.Execute(Me.commonControl.Name)
    If Not MatchCollection.Count = 1 Then
        '検索結果が1件じゃない場合
        MsgBox "コントロール名の検索に失敗しました Day Label Click"
        GoTo CloseAndExit
    End If
    Dim dateTarget As Date
    'フォームの日付配列よりクリックされた日付を特定し、取得
    dateTarget = frmDatePicker.vardateDayArray_CurrentYearMonth(CLng(MatchCollection(0).SubMatches(1)))
    '取得したDateをFormCommonのグローバル変数にセット
    FormCommon.datePickerResult = dateTarget
    'DatePickerフォームをUnload
    Unload frmDatePicker
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "common_DatePicker_Label_Click code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
'''DatePicker cmbBox
'''年月、1日として、Date型になるかどうかチェックし、日付ラベルの設定まで行う
Private Sub common_DatePicker_cmbBox_Change()
    On Error GoTo ErrorCatch
    If Not IsNumeric(frmDatePicker.cmbBox_Year.Text) Or Not IsNumeric(frmDatePicker.cmdBox_Month.Text) Then
        MsgBox "年月どちらかの値が数値ととして認識できませんでした"
        GoTo CloseAndExit
        Exit Sub
    End If
    If Not IsDate(DateSerial(CInt(frmDatePicker.cmbBox_Year.Text), CInt(frmDatePicker.cmdBox_Month.Text), 1)) Then
        '年月1日の組み合わせがDateとして認識できなかった
        MsgBox "設定された年月の組み合わせが日付として認識できませんでした"
        GoTo CloseAndExit
    End If
    '日付ラベルの数分のDate配列を用意
    Dim datearrDay(41) As Date
    Dim date1dayCurYM As Date           '年月1日のDate
    date1dayCurYM = DateSerial(CInt(frmDatePicker.cmbBox_Year.Text), CInt(frmDatePicker.cmdBox_Month.Text), 1)
    'まずは当月の分の日付をセットしてやる
    Dim longArrayRowCounter As Long
    'WeekDayが1(日曜)から始まることを利用し、WeekDay() -2を得て、後に配列の添え字からこの数字を引けば日付のオフセットが得られる
    '配列の添え字のエンドは翌月0日(当月末の日にち)
    '上記2つをそれぞれ算出する
    Dim longStartDayWeekday As Long
    longStartDayWeekday = Weekday(date1dayCurYM) - 2
    Dim longEndRow As Long
    longEndRow = Day(DateSerial(Year(date1dayCurYM), Month(date1dayCurYM) + 1, 0))
    For longArrayRowCounter = LBound(datearrDay) To UBound(datearrDay)
        datearrDay(longArrayRowCounter) = DateSerial(Year(date1dayCurYM), Month(date1dayCurYM), longArrayRowCounter - longStartDayWeekday)
    Next longArrayRowCounter
    frmDatePicker.vardateDayArray_CurrentYearMonth = datearrDay
    '日付ラベルにセットしていく
    'DatePickerのすべてのコントロールをループし、日付ラベルの命名規則に沿ったコントロールのみ操作する
#If refRegExp Then
    'RegExp参照済み
    Dim objRegExp As RegExp
    Set objRegExp = New RegExp
#Else
    'RegExp参照なし
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
#End If
    '正規表現パターン設定
    'プリフィックスと数字部分2グループに分ける
    objRegExp.Pattern = "(" & FormCommon.LABEL_DAY_PRIFIX & ")([0-9]{2})"
    Dim ContrlEach As MSForms.Control
    For Each ContrlEach In frmDatePicker.Controls
        If objRegExp.Test(ContrlEach.Name) Then
#If refRegExp Then
            'regexp参照されている場合
            Dim MatchCollection As MatchCollection
#Else
            '参照無しの場合
            Dim MatchCollection As Object
#End If
            Set MatchCollection = objRegExp.Execute(ContrlEach.Name)
            If Not MatchCollection.Count = 1 Then
                MsgBox "コントロールの名前パターンの検索に失敗しました"
                GoTo CloseAndExit
            End If
            'MatchCollection(0).submatches(1)で検索結果1番目の2要素目(数字グループ)の結果が返る
            ContrlEach.Caption = Day(datearrDay(CLng(MatchCollection(0).SubMatches(1))))
            '日付の条件により背景の色や文字色を変える
            Select Case True
            Case datearrDay(CLng(MatchCollection(0).SubMatches(1))) = DateSerial(Year(Now()), Month(Now()), Day(Now()))
                '当日にぴったりだったら
                '文字色を標準黒
                ContrlEach.ForeColor = FormCommon.TEXT_NORMAL_BLACK
                '背景色を薄緑色に
                ContrlEach.BackColor = FormCommon.TXTBOX_BACKCOLORE_EDITABLE
            Case Month(datearrDay(CLng(MatchCollection(0).SubMatches(1)))) = Month(date1dayCurYM)
                '当月の日にちだった場合、文字色をノーマルに
                ContrlEach.ForeColor = FormCommon.TEXT_NORMAL_BLACK
                '土日だけは色を変える、それ以外はピンク
                Select Case Weekday(datearrDay(CLng(MatchCollection(0).SubMatches(1))))
                Case vbSaturday
                    '土曜、背景青っぽいのに
                    ContrlEach.BackColor = FormCommon.BACK_COLOR_SATURDAY_BLUE
                Case vbSunday
                    '日曜、背景朱色っぽいのに
                    ContrlEach.BackColor = FormCommon.BACK_COLOR_SUNDAY_RED
                Case Else
                    'それ以外
                    '背景色は選択可能ピンクへ
                    ContrlEach.BackColor = FormCommon.BACK_COLOR_SELECTABLE_PINK
                End Select
            Case Month(datearrDay(CLng(MatchCollection(0).SubMatches(1)))) <> Month(date1dayCurYM)
                '当月の日にちじゃなかったら薄い灰色に
                ContrlEach.ForeColor = FormCommon.TEXT_NOT_ENABLE_GRAY
                '曜日によって背景色変更
                Select Case Weekday(datearrDay(CLng(MatchCollection(0).SubMatches(1))))
                Case vbSaturday
                    '土曜、背景青っぽいのに
                    ContrlEach.BackColor = FormCommon.BACK_COLOR_SATURDAY_BLUE
                Case vbSunday
                    '日曜、背景朱色っぽいのに
                    ContrlEach.BackColor = FormCommon.BACK_COLOR_SUNDAY_RED
                Case Else
                    'それ以外
                    '背景色は選択可能ピンクへ
                    ContrlEach.BackColor = FormCommon.BACK_COLOR_SELECTABLE_PINK
                End Select
            End Select
        End If
        Set MatchCollection = Nothing
    Next ContrlEach
    GoTo CloseAndExit
ErrorCatch:
    DebugMsgWithTime "common_DatePicker_cmbBox_Change code: " & Err.Number & " Description: " & Err.Description
    GoTo CloseAndExit
CloseAndExit:
    Exit Sub
End Sub
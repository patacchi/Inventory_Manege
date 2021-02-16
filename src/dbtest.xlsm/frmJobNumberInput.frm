VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJobNumberInput 
   Caption         =   "ジョブ番号・履歴登録画面"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7395
   OleObjectBlob   =   "frmJobNumberInput.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmJobNumberInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Option Base 1
Private KishuInfoInfrmJobInput As typKishuInfo
Private Sub frmJobNumberInput_Initialize()
    'フォーム初期化（全部消すだけ）
    txtboxJobNumber.Text = ""
    txtboxMaisuu.Text = ""
    txtboxStartRireki = ""
    labelZuban.Caption = ""
    strQRZuban = ""
    strRegistRireki = ""
    btnQRFormShow.SetFocus
End Sub
Private Sub btnInputRirekiNumber_Click()
    'ジョブ番号・履歴の登録処理
    'Dim KishuInfoInfrmJobInput As typKishuInfo
    Dim isCollect As Boolean
    Dim sqlbJobInput As clsSQLStringBuilder
    Dim dblTimer As Double
    Dim longStartRirekiNumber As Long
    Dim longEndRirekiNumber As Long
    Dim longDuplicateNumber As Long                 '重複履歴の数
    Dim longMsgBoxRetCode As Long
    On Error GoTo ErrorCatch
    If txtboxJobNumber.Text = Empty Or _
        txtboxMaisuu.Text = Empty Or _
        txtboxStartRireki.Text = Empty Then
        MsgBox ("空白の項目があります。確認してください")
        Exit Sub
    End If
    If CLng(txtboxMaisuu.Text) < 1 Then
        MsgBox ("枚数には1以上の整数を入力して下さい")
        txtboxMaisuu.SetFocus
        Exit Sub
    End If
    '時間計測開始
    If Not labelZuban.Caption = "" Then
        strQRZuban = labelZuban.Caption
    End If
    dblTimer = timer()
    'スタート履歴からKishuInfoを引っ張ってくる
    KishuInfoInfrmJobInput = getKishuInfoByRireki(txtboxStartRireki.Text)
    If KishuInfoInfrmJobInput.KishuHeader = "" Then
        'KishuInfoが取れてないということは失敗してるっぽいので、ここで処理中止
        Exit Sub
    End If
    Set sqlbJobInput = New clsSQLStringBuilder
    With sqlbJobInput
        .JobNumber = CStr(txtboxJobNumber.Text)
        .FieldArray = arrFieldList_JobData
        .StartRireki = CStr(txtboxStartRireki.Text)
        .Maisu = CLng(txtboxMaisuu.Text)
        .RenbanKeta = KishuInfoInfrmJobInput.RenbanKetasuu
        .TableName = Table_JobDataPri & KishuInfoInfrmJobInput.KishuName
    End With
    Set sqlbJobInput.FieldType = GetFieldTypeNameByTableName(sqlbJobInput.TableName)
    If Not Len(txtboxStartRireki.Text) = KishuInfoInfrmJobInput.TotalRirekiketa Then
        MsgBox "履歴の桁数が登録されている機種名：" & KishuInfoInfrmJobInput.KishuName & " の " & _
                KishuInfoInfrmJobInput.TotalRirekiketa & " 桁と違います。処理を中止します。"
                txtboxStartRireki.SetFocus
                GoTo CloseAndExit
    End If
    'スタート履歴（の連番）とエンド履歴を算出し、重複がないかチェック
    longStartRirekiNumber = CLng(Right(sqlbJobInput.StartRireki, KishuInfoInfrmJobInput.RenbanKetasuu))
    longEndRirekiNumber = longStartRirekiNumber + sqlbJobInput.Maisu - 1
    longDuplicateNumber = GetRecordCountSimple(sqlbJobInput.TableName, Job_RirekiNumber, _
                            "BETWEEN " & longStartRirekiNumber & " AND " & longEndRirekiNumber & ";")
    If longDuplicateNumber >= 1 Then
        '重複があったらしい
        longMsgBoxRetCode = MsgBox(prompt:="登録しようとしているデータで " & longDuplicateNumber & " 件の重複があったようです。入力しなおしますか？" _
                            , Buttons:=vbYesNo)
        Select Case longMsgBoxRetCode
        Case vbYes
            '入力しなおし、つまり何もしないで脱出
            txtboxStartRireki.SetFocus
            GoTo CloseAndExit
        Case vbNo
            'そのまま続行
        End Select
    End If
    isCollect = sqlbJobInput.CreateInsertSQL(boolCheckLastRireki:=True)
    If Not isCollect Then
        MsgBox "ジョブ登録中に何かあったようです"
        GoTo CloseAndExit
        Exit Sub
    End If
    MsgBox "ジョブ登録完了。 " & sqlbJobInput.Maisu & " 枚のデータを " & timer() - dblTimer & " 秒で処理しました"
    Debug.Print "ジョブ登録完了。 " & sqlbJobInput.Maisu & " 枚のデータを " & timer() - dblTimer & " 秒で処理しました"
    frmJobNumberInput_Initialize
    GoTo CloseAndExit
    Exit Sub
CloseAndExit:
    Set sqlbJobInput = Nothing
    Exit Sub
ErrorCatch:
    Debug.Print "ImputRireki Erro code: " & Err.Number & "Description: " & Err.Description
    Set sqlbJobInput = Nothing
    Exit Sub
End Sub
Private Sub btnQRFormShow_Click()
    Dim KishuLocal As typKishuInfo
    Dim varReturn As Variant
    Dim strNewRireki As String
    'QRコード読み取りフォーム表示
'    frmJobNumberInput.Hide
    frmQRAnalyze.Show
    If Not labelZuban.Caption = "" Then
        '図番に何か入ってたらKishuInfoを取得してみて、その情報をもとに機種ヘッダを履歴ボックスに入力してやる
        KishuLocal = GetKishuinfoByZuban(labelZuban.Caption)
        If KishuLocal.KishuHeader = "" Then
            Debug.Print "QRコードからKishuInfo引っ張ったけど空だった"
            txtboxStartRireki.Text = ""
            Exit Sub
        End If
        '最新情報入力チェックボックスがTrueの場合は最新履歴を入力してやる
        If chkboxInputNextNumber = True Then
            '最新履歴を取得
            strNewRireki = GetNextRireki(Table_JobDataPri & KishuLocal.KishuName)
            If strNewRireki = "" Then
                Debug.Print "最新履歴取得したが、空だった"
                '空だったらヘッダは入れてやろう
                txtboxStartRireki.Text = KishuLocal.KishuHeader
                Exit Sub
            End If
            txtboxStartRireki.Text = strNewRireki
        Else
            'チェックされてない場合は、ヘッダのみを入力する
            txtboxStartRireki.Text = KishuLocal.KishuHeader
        End If
    End If
End Sub
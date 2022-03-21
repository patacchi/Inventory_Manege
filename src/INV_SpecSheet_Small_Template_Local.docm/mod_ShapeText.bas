Attribute VB_Name = "mod_ShapeText"
Option Explicit
'Rangeオブジェクトの動作が重い事への対処、指定文字数未満は無視する
Private Const MIN_SHAPE_CHAR_LENGTHB As Long = 9                    'フォントサイズ調整時、この定数以上のバイト数の文字列を整形対象とする、range.Informationの動作が遅いため対象を少なく
Private Const MIN_FONT_SIZE As Long = 5                             '最小フォントサイズ、このフォントになったら縮小処理を中断する
Private Const SCALE_MESURE_API_TO_REAL As Single = 0.73!            'APIのMesureTextWidthが何故か大きい数値を返すので、係数を掛ける
'''2行にまたがる段落を１行に収まるようにフォントサイズを調整する
'''
Public Sub CalcFontSizeToOneLine()
    '画面更新を停止する
    Application.ScreenUpdating = False
    Dim strarrPerParagraph() As String
    'ActiveDocumentの全文字を取得し、段落記号(文字コード 13)で分割し、配列に収める
    'この配列の添え字 +1 がParagrafs(x)のIndex番号と一致する(Paragrafsはインデックス1スタート)
    strarrPerParagraph = Split(Replace(ActiveDocument.Range(0, ActiveDocument.Range.End).Text, ChrW(7), ""), ChrW(13))
    '取得した段落配列をループ
    Dim longParaArrayRowCounter As Long
    '配列の最後は空文字なので無視する
    For longParaArrayRowCounter = LBound(strarrPerParagraph) To UBound(strarrPerParagraph) - 1
        Select Case True
        Case LenB(strarrPerParagraph(longParaArrayRowCounter)) > MIN_SHAPE_CHAR_LENGTHB
            '対象最小バイト数を上回っていた場合
            '行開始位置調査、フォントサイズ設定対象
            '最後の改行抜きのRangeを取得
            Dim rngNoCrlf As Range
            '段落配列の添え字+1がParagrafsのIndex
            Set rngNoCrlf = ActiveDocument.Range( _
                            ActiveDocument.Paragraphs(longParaArrayRowCounter + 1).Range.Start, _
                            ActiveDocument.Paragraphs(longParaArrayRowCounter + 1).Range.End - 1 _
                            )
            '開始位置と終了位置を取得(行数だとだめだった)
            '横拡大掛けるとうまくいかないのでAPIのを使ってみよう・・・
            Dim longPixColumnWidth As Long          '現在の列の幅のPixcelを取得
            Dim longPixCurrent As Long              'カレント文字列のPicelを取得
            '現在の列の幅のPixcelを取得
            longPixColumnWidth = Application.PointsToPixels(rngNoCrlf.Cells.Width)
            'カレント文字列の幅を取得
            longPixCurrent = CSng(Mod_WinAPI.MesureTextWidth(rngNoCrlf.Text, rngNoCrlf.Font.Name, rngNoCrlf.Font.Size, rngNoCrlf.Font.Scaling)) * SCALE_MESURE_API_TO_REAL
            'カレント文字列がセルの幅より長い間ループする
            Do While longPixCurrent > longPixColumnWidth
                If rngNoCrlf.Font.Size <= MIN_FONT_SIZE Then
                    '最小フォントサイズより小さかった
                    MsgBox "表示が2行に分かれている可能性がありますが、文字をこれ以上小さくできないため処理を中断しました" & vbCrLf & _
                    "対象文字列：" & strarrPerParagraph(longParaArrayRowCounter)
                End If
                'FontSizeを -0.5
                rngNoCrlf.Font.Size = rngNoCrlf.Font.Size - 0.5
                '再度幅を取得
                '現在の列の幅のPixcelを取得
                longPixColumnWidth = Application.PointsToPixels(rngNoCrlf.Cells.Width)
                'カレント文字列の幅を取得
                longPixCurrent = Mod_WinAPI.MesureTextWidth(rngNoCrlf.Text, rngNoCrlf.Font.Name, rngNoCrlf.Font.Size, rngNoCrlf.Font.Scaling)
            Loop
'            Dim sglStartVPos As Single       '段落の最初の文字の垂直方向の上端からのポイント数
'            Dim sglEndVPos As Single         '段落の最後の文字の垂直方向の上端からのポイント数
'            sglStartVPos = rngNoCrlf.Information(wdVerticalPositionRelativeToPage)
'            sglEndVPos = rngNoCrlf.Characters.Last.Information(wdVerticalPositionRelativeToPage)
'            スタート行とエンド行が違う間はループする
'            Do While sglStartVPos <> sglEndVPos
'                スタート行とEnd行が違った
'                If rngNoCrlf.Font.Size <= MIN_FONT_SIZE Then
'                    最小フォントサイズより小さかった
'                    MsgBox "表示が2行に分かれている可能性がありますが、文字をこれ以上小さくできないため処理を中断しました" & vbCrLf & _
'                    "対象文字列：" & strarrPerParagraph(longParaArrayRowCounter)
'                End If
'                FontSizeを -0.5
'                rngNoCrlf.Font.Size = rngNoCrlf.Font.Size - 0.5
'                再度開始行と終了行を取得する
'                sglStartVPos = rngNoCrlf.Information(wdVerticalPositionRelativeToPage)
'                sglEndVPos = rngNoCrlf.Characters.Last.Information(wdVerticalPositionRelativeToPage)
'            Loop
'            Debug.Print longParaArrayRowCounter & vbCrLf & "Text: " & rngNoCrlf.Text & vbCrLf & "StartLine: " & sglStartVPos & vbCrLf & "End Line: " & sglEndVPos
            Set rngNoCrlf = Nothing
            '制御をWindowsに戻す
            DoEvents
        End Select  'LenB
    Next longParaArrayRowCounter
    '画面更新を再開する
    Application.ScreenUpdating = True
End Sub
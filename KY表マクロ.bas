Attribute VB_Name = "KY表マクロ"
Option Explicit

'==============================================
' KY表 自動作成マクロ  Ver 1.2
' 公民連携沖縄株式会社
'==============================================

Private Const SHEET_INPUT As String = "入力"
Private Const SHEET_KY    As String = "KY表"
Private Const COMPANY     As String = "公民連携沖縄株式会社"
Private Const REIWA_BASE  As Integer = 2018

'----------------------------------------------
' 初回セットアップ（一度だけ実行）
'----------------------------------------------
Public Sub セットアップ()
    Application.ScreenUpdating = False
    Call DeleteSheet(SHEET_INPUT)
    Call DeleteSheet(SHEET_KY)
    Call CreateKYSheet
    Call CreateInputSheet
    Application.ScreenUpdating = True
    MsgBox "セットアップ完了！" & vbCrLf & _
           "「" & SHEET_INPUT & "」シートに記入後、ボタンを押してください。", _
           vbInformation, "KY表マクロ"
End Sub

'----------------------------------------------
' KY表を作成して印刷プレビューを開く
'----------------------------------------------
Public Sub KY表作成()
    Dim wsIn As Worksheet
    Dim wsKY As Worksheet

    On Error Resume Next
    Set wsIn = ThisWorkbook.Sheets(SHEET_INPUT)
    Set wsKY = ThisWorkbook.Sheets(SHEET_KY)
    On Error GoTo 0

    If wsIn Is Nothing Or wsKY Is Nothing Then
        MsgBox "先に「セットアップ」マクロを実行してください。", vbExclamation
        Exit Sub
    End If

    Dim m As String, d As String
    m = Trim(CStr(wsIn.Range("C3").Value))
    d = Trim(CStr(wsIn.Range("C4").Value))

    If m = "" Then MsgBox "「月」を入力してください。", vbExclamation : Exit Sub
    If d = "" Then MsgBox "「日」を入力してください。", vbExclamation : Exit Sub

    Dim yobi As String
    yobi = GetYobi(CInt(m), CInt(d))

    Application.ScreenUpdating = False

    ' クリア
    wsKY.Range("F2").Value = ""
    wsKY.Range("H2").Value = ""
    wsKY.Range("J2").Value = ""
    wsKY.Range("L2").Value = ""
    wsKY.Range("C3").Value = ""
    wsKY.Range("A5").Value = ""
    wsKY.Range("H5").Value = ""
    wsKY.Range("A7").Value = ""
    wsKY.Range("H7").Value = ""
    wsKY.Range("D11").Value = ""
    wsKY.Range("M13").Value = ""
    wsKY.Range("B14").Value = ""
    wsKY.Range("M14").Value = ""

    ' 転記
    wsKY.Range("F2").Value = m                            ' 月
    wsKY.Range("H2").Value = d                            ' 日
    wsKY.Range("J2").Value = yobi                         ' 曜日（自動）
    wsKY.Range("L2").Value = wsIn.Range("C5").Value       ' 天候
    wsKY.Range("C3").Value = wsIn.Range("C6").Value       ' グループの作業内容
    wsKY.Range("A5").Value = wsIn.Range("C7").Value       ' 危険のポイント①
    wsKY.Range("H5").Value = wsIn.Range("C8").Value       ' 私達はこうする①
    wsKY.Range("A7").Value = wsIn.Range("C9").Value       ' 危険のポイント②
    wsKY.Range("H7").Value = wsIn.Range("C10").Value      ' 私達はこうする②
    wsKY.Range("D11").Value = wsIn.Range("C11").Value     ' 本日の安全目標
    wsKY.Range("M13").Value = wsIn.Range("C12").Value     ' リーダー名
    wsKY.Range("B14").Value = wsIn.Range("C13").Value     ' 参加者氏名
    wsKY.Range("M14").Value = wsIn.Range("C14").Value     ' 作業員数

    Application.ScreenUpdating = True
    wsKY.Activate
    wsKY.PrintPreview
    wsIn.Activate
End Sub

'----------------------------------------------
' 曜日取得
'----------------------------------------------
Private Function GetYobi(m As Integer, d As Integer) As String
    Dim dt As Date
    Dim arr As Variant
    arr = Array("日", "月", "火", "水", "木", "金", "土")
    On Error GoTo ErrExit
    dt = DateSerial(Year(Now), m, d)
    GetYobi = arr(Weekday(dt) - 1)
    Exit Function
ErrExit:
    GetYobi = "　"
End Function

'----------------------------------------------
' シート削除
'----------------------------------------------
Private Sub DeleteSheet(sName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

'==============================================================
' 入力シート作成
'==============================================================
Private Sub CreateInputSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SHEET_INPUT

    With ws.Range("A1:E1")
        .Merge
        .Value = "KY表　入力フォーム"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 70, 127)
    End With
    ws.Rows(1).RowHeight = 32

    With ws.Range("A2:E2")
        .Merge
        .Value = "※ 黄色セルに入力後、「KY表を作成」ボタンを押してください"
        .Font.Size = 9
        .Font.Color = RGB(150, 150, 150)
    End With
    ws.Rows(2).RowHeight = 16

    Dim labels(11) As String
    Dim heights(11) As Double

    labels(0)  = "月"
    labels(1)  = "日"
    labels(2)  = "天候"
    labels(3)  = "グループの作業内容"
    labels(4)  = "危険のポイント ①"
    labels(5)  = "私達はこうする ①"
    labels(6)  = "危険のポイント ②"
    labels(7)  = "私達はこうする ②"
    labels(8)  = "本日の安全目標"
    labels(9)  = "リーダー名"
    labels(10) = "作業員数（名）"
    labels(11) = "参加者氏名"

    heights(0)  = 20 : heights(1)  = 20 : heights(2)  = 20
    heights(3)  = 50 : heights(4)  = 50 : heights(5)  = 50
    heights(6)  = 50 : heights(7)  = 50 : heights(8)  = 32
    heights(9)  = 20 : heights(10) = 20 : heights(11) = 50

    Dim i As Integer
    For i = 0 To 11
        Dim r As Integer
        r = i + 3
        ws.Rows(r).RowHeight = heights(i)

        With ws.Range("B" & r)
            .Value = labels(i)
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignRight
            .VerticalAlignment = xlVAlignCenter
        End With

        With ws.Range("C" & r)
            .Interior.Color = RGB(255, 255, 180)
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
        End With

        ws.Range("B" & r).BorderAround xlContinuous, xlThin
        ws.Range("C" & r).BorderAround xlContinuous, xlThin
    Next i

    ws.Range("C5").Value = "晴れ"
    With ws.Range("D5")
        .Value = "← 晴れ・曇り・雨・荒天"
        .Font.Color = RGB(150, 150, 150)
        .Font.Size = 9
        .VerticalAlignment = xlVAlignCenter
    End With

    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("D").ColumnWidth = 22

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("C16").Left, ws.Range("C16").Top + 8, 180, 36)
    With shp
        .Name = "btnKY"
        .TextFrame.Characters.Text = "▶  KY表を作成する"
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Line.Visible = msoFalse
        .OnAction = "KY表作成"
    End With

    ws.Range("C3").Select
End Sub

'==============================================================
' KY表テンプレートシート作成（ユーザー指定フォント・列幅）
'==============================================================
Private Sub CreateKYSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_KY

    ' ページ設定（A4横）
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin    = Application.CentimetersToPoints(1.0)
        .BottomMargin = Application.CentimetersToPoints(1.0)
        .LeftMargin   = Application.CentimetersToPoints(1.5)
        .RightMargin  = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
    End With

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' ---- 列幅（ユーザー指定値）----
    ws.Columns("A").ColumnWidth = 5.75
    ws.Columns("B").ColumnWidth = 8.25
    ws.Columns("C").ColumnWidth = 8.25
    ws.Columns("D").ColumnWidth = 8.25
    ws.Columns("E").ColumnWidth = 8.25
    ws.Columns("F").ColumnWidth = 8.25
    ws.Columns("G").ColumnWidth = 9.88
    ws.Columns("H").ColumnWidth = 8.25
    ws.Columns("I").ColumnWidth = 8.25
    ws.Columns("J").ColumnWidth = 8.25
    ws.Columns("K").ColumnWidth = 8.25
    ws.Columns("L").ColumnWidth = 26.5
    ws.Columns("M").ColumnWidth = 10.3
    ws.Columns("N").ColumnWidth = 5.75

    ' ---- 行高 ----
    ws.Rows("1").RowHeight  = 4
    ws.Rows("2").RowHeight  = 52   ' タイトル・日付行
    ws.Rows("3").RowHeight  = 28   ' 作業内容
    ws.Rows("4").RowHeight  = 26   ' KYヘッダー
    ws.Rows("5").RowHeight  = 44   ' KY①上
    ws.Rows("6").RowHeight  = 44   ' KY①下
    ws.Rows("7").RowHeight  = 44   ' KY②上
    ws.Rows("8").RowHeight  = 44   ' KY②下
    ws.Rows("9").RowHeight  = 44   ' KY③上（空白）
    ws.Rows("10").RowHeight = 44   ' KY③下（空白）
    ws.Rows("11").RowHeight = 28   ' 安全目標
    ws.Rows("12").RowHeight = 34   ' チェック項目
    ws.Rows("13").RowHeight = 26   ' 会社名・リーダー
    ws.Rows("14").RowHeight = 30   ' 参加者・作業員

    ' ==== 行2: タイトル＋日付 ====
    With ws.Range("A2:D2")
        .Merge
        .Value = "危険予知活動表"
        .Font.Size = 48
        .Font.Bold = True
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
    End With

    Dim reiwaYear As Integer
    reiwaYear = Year(Now) - REIWA_BASE

    ws.Range("E2").Value = "令和" & reiwaYear & "年"
    ws.Range("F2").Value = ""           ' ← 月（データ）
    ws.Range("G2").Value = "月"
    ws.Range("H2").Value = ""           ' ← 日（データ）
    ws.Range("I2").Value = "日（"
    ws.Range("J2").Value = ""           ' ← 曜日（自動）
    ws.Range("K2").Value = "曜日）天候"
    ws.Range("L2").Value = ""           ' ← 天候（データ）

    Dim cl As Range
    For Each cl In ws.Range("E2:L2")
        cl.Font.Size = 18
        cl.VerticalAlignment = xlVAlignBottom
        cl.HorizontalAlignment = xlHAlignCenter
    Next cl

    ws.Range("F2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("H2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("J2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("L2").Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' ==== 行3: グループの作業内容 ====
    With ws.Range("A3:B3")
        .Merge
        .Value = "グループの作業内容"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("C3:N3")
        .Merge
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .Font.Size = 18
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行4: KYヘッダー ====
    With ws.Range("A4:G4")
        .Merge
        .Value = "危険のポイント"
        .Font.Bold = True
        .Font.Size = 28
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("H4:N4")
        .Merge
        .Value = "私達はこうする"
        .Font.Bold = True
        .Font.Size = 28
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行5-10: KYデータ（2行1セット × 3セット）====
    Dim startRow As Integer
    For startRow = 5 To 9 Step 2
        With ws.Range("A" & startRow & ":G" & (startRow + 1))
            .Merge
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
            .Font.Size = 16
            .BorderAround xlContinuous, xlMedium
        End With
        With ws.Range("H" & startRow & ":N" & (startRow + 1))
            .Merge
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
            .Font.Size = 16
            .BorderAround xlContinuous, xlMedium
        End With
    Next startRow

    ' ==== 行11: 本日の安全目標 ====
    With ws.Range("A11:C11")
        .Merge
        .Value = "本日の安全目標"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("D11:N11")
        .Merge
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .Font.Size = 20
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行12: チェック項目 ====
    With ws.Range("A12:N12")
        .Merge
        .Value = "□ 現地の安全確認は実施したか　　□ 体調不良者はいないか（無し・対処済）　　□ 服装・装備は万全か" & Chr(10) & _
                 "□ 天候や視界に問題はないか（無し・対処済）　　□ 事前の予定と異なった要素はないか（無し・対処済）"
        .Font.Size = 14
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行13: 会社名・リーダー ====
    With ws.Range("A13")
        .Value = "会社名"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("B13:K13")
        .Merge
        .Value = COMPANY
        .Font.Size = 22
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("L13")
        .Value = "リーダー"
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("M13:N13")
        .Merge
        .Font.Size = 16
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行14: 参加者・作業員 ====
    With ws.Range("A14")
        .Value = "参加者" & Chr(10) & "サイン"
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("B14:K14")
        .Merge
        .Font.Size = 18
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("L14")
        .Value = "作業員"
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("M14")
        .Font.Size = 18
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("N14")
        .Value = "名"
        .Font.Bold = True
        .Font.Size = 26
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With
End Sub

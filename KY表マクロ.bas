Attribute VB_Name = "KY表マクロ"
Option Explicit

'==============================================
' KY表 自動作成マクロ  Ver 1.0
' 公民連携沖縄株式会社
'==============================================
'
' 【使い方】
' 1. このコードをExcelのVBAエディタ（Alt+F11）→ Module1 に貼り付ける
' 2. マクロ「セットアップ」を一度だけ実行する
' 3. 「入力」シートの黄色セルに記入する
' 4. 「KY表を作成する」ボタンを押す → 印刷プレビューが開く
'==============================================

Private Const SHEET_INPUT As String = "入力"
Private Const SHEET_KY    As String = "KY表"
Private Const COMPANY     As String = "公民連携沖縄株式会社"
Private Const REIWA_BASE  As Integer = 2018  ' 令和元年=2019、2018を引いて算出

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

    ' 必須チェック
    Dim m As String, d As String
    m = Trim(CStr(wsIn.Range("C3").Value))
    d = Trim(CStr(wsIn.Range("C4").Value))

    If m = "" Then
        MsgBox "「月」を入力してください。", vbExclamation : Exit Sub
    End If
    If d = "" Then
        MsgBox "「日」を入力してください。", vbExclamation : Exit Sub
    End If

    ' 曜日自動計算
    Dim yobi As String
    yobi = GetYobi(CInt(m), CInt(d))

    ' KY表をクリア（前回データ削除）
    Application.ScreenUpdating = False
    wsKY.Range("G2,I2,K2,N2").ClearContents
    wsKY.Range("D3").ClearContents
    wsKY.Range("A5,I5,A7,I7").ClearContents
    wsKY.Range("D11,M13,C14,M14").ClearContents

    ' データ転記
    wsKY.Range("G2").Value = m                            ' 月
    wsKY.Range("I2").Value = d                            ' 日
    wsKY.Range("K2").Value = yobi                         ' 曜日（自動）
    wsKY.Range("N2").Value = wsIn.Range("C5").Value       ' 天候
    wsKY.Range("D3").Value = wsIn.Range("C6").Value       ' グループの作業内容
    wsKY.Range("A5").Value = wsIn.Range("C7").Value       ' 危険のポイント①
    wsKY.Range("I5").Value = wsIn.Range("C8").Value       ' 私達はこうする①
    wsKY.Range("A7").Value = wsIn.Range("C9").Value       ' 危険のポイント②
    wsKY.Range("I7").Value = wsIn.Range("C10").Value      ' 私達はこうする②
    wsKY.Range("D11").Value = wsIn.Range("C11").Value     ' 本日の安全目標
    wsKY.Range("M13").Value = wsIn.Range("C12").Value     ' リーダー名
    wsKY.Range("C14").Value = wsIn.Range("C13").Value     ' 参加者氏名
    wsKY.Range("M14").Value = wsIn.Range("C14").Value     ' 作業員数

    Application.ScreenUpdating = True

    ' KY表シートで印刷プレビュー
    wsKY.Activate
    wsKY.PrintPreview
    wsIn.Activate
End Sub

'----------------------------------------------
' 曜日取得（日・月・火・水・木・金・土）
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
' シート削除ユーティリティ
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

    ' タイトル
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

    ' 入力項目定義
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

    ' 天候デフォルト＆ヒント
    ws.Range("C5").Value = "晴れ"
    With ws.Range("D5")
        .Value = "← 晴れ・曇り・雨・荒天"
        .Font.Color = RGB(150, 150, 150)
        .Font.Size = 9
        .VerticalAlignment = xlVAlignCenter
    End With

    ' 列幅
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("D").ColumnWidth = 22

    ' 実行ボタン
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
' KY表テンプレートシート作成
'==============================================================
Private Sub CreateKYSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_KY

    ' ページ設定（A4横）
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin    = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.2)
        .LeftMargin   = Application.CentimetersToPoints(1.5)
        .RightMargin  = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
    End With

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' 列幅（A-P: 16列でA4横）
    ws.Columns("A:B").ColumnWidth = 5.5
    ws.Columns("C:H").ColumnWidth = 8.5
    ws.Columns("I:P").ColumnWidth = 8.5

    ' 行高
    ws.Rows("1").RowHeight  = 4
    ws.Rows("2").RowHeight  = 40  ' タイトル＋日付
    ws.Rows("3").RowHeight  = 30  ' 作業内容
    ws.Rows("4").RowHeight  = 22  ' KYヘッダー
    ws.Rows("5").RowHeight  = 22  ' KY①上
    ws.Rows("6").RowHeight  = 22  ' KY①下
    ws.Rows("7").RowHeight  = 22  ' KY②上
    ws.Rows("8").RowHeight  = 22  ' KY②下
    ws.Rows("9").RowHeight  = 22  ' KY③上（空白）
    ws.Rows("10").RowHeight = 22  ' KY③下（空白）
    ws.Rows("11").RowHeight = 26  ' 安全目標
    ws.Rows("12").RowHeight = 32  ' チェック項目
    ws.Rows("13").RowHeight = 24  ' 会社名・リーダー
    ws.Rows("14").RowHeight = 28  ' 参加者・作業員

    ' ==== 行2: タイトル + 日付 ====
    With ws.Range("A2:E2")
        .Merge
        .Value = "危険予知活動表"
        .Font.Size = 26
        .Font.Bold = True
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
        .IndentLevel = 1
    End With

    ' 令和年（自動計算）
    Dim reiwaYear As Integer
    reiwaYear = Year(Now) - REIWA_BASE

    ws.Range("F2").Value = "令和" & reiwaYear & "年"
    ws.Range("G2").Value = ""       ' ← 月（データ入力）
    ws.Range("H2").Value = "月"
    ws.Range("I2").Value = ""       ' ← 日（データ入力）
    ws.Range("J2").Value = "日（"
    ws.Range("K2").Value = ""       ' ← 曜日（自動）
    ws.Range("L2").Value = "曜日）"
    ws.Range("M2").Value = "天候"
    ws.Range("N2").Value = ""       ' ← 天候（データ入力）

    Dim cl As Range
    For Each cl In ws.Range("F2:N2")
        cl.Font.Size = 10
        cl.VerticalAlignment = xlVAlignBottom
        cl.HorizontalAlignment = xlHAlignCenter
    Next cl

    ws.Range("G2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("I2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("K2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("N2").Borders(xlEdgeBottom).LineStyle = xlContinuous

    ' ==== 行3: グループの作業内容 ====
    With ws.Range("A3:C3")
        .Merge
        .Value = "グループの作業内容"
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("D3:P3")
        .Merge
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .Font.Size = 11
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行4: KYヘッダー ====
    With ws.Range("A4:H4")
        .Merge
        .Value = "危険のポイント"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("I4:P4")
        .Merge
        .Value = "私達はこうする"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行5-10: KYデータ（2行1セット × 3セット）====
    Dim startRow As Integer
    For startRow = 5 To 9 Step 2
        With ws.Range("A" & startRow & ":H" & (startRow + 1))
            .Merge
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
            .Font.Size = 11
            .BorderAround xlContinuous, xlMedium
        End With
        With ws.Range("I" & startRow & ":P" & (startRow + 1))
            .Merge
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
            .Font.Size = 11
            .BorderAround xlContinuous, xlMedium
        End With
    Next startRow

    ' ==== 行11: 本日の安全目標 ====
    With ws.Range("A11:C11")
        .Merge
        .Value = "本日の安全目標"
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = RGB(220, 220, 220)
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("D11:P11")
        .Merge
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .Font.Size = 11
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行12: チェック項目 ====
    With ws.Range("A12:P12")
        .Merge
        .Value = "□ 現地の安全確認は実施したか　　□ 体調不良者はいないか（無し・対処済）　　□ 服装・装備は万全か" & Chr(10) & _
                 "□ 天候や視界に問題はないか（無し・対処済）　　□ 事前の予定と異なった要素はないか（無し・対処済）"
        .Font.Size = 9
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行13: 会社名・リーダー ====
    With ws.Range("A13:B13")
        .Merge
        .Value = "会社名"
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("C13:K13")
        .Merge
        .Value = COMPANY
        .Font.Size = 11
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("L13")
        .Value = "リーダー"
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("M13:P13")
        .Merge
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    ' ==== 行14: 参加者・作業員 ====
    With ws.Range("A14:B14")
        .Merge
        .Value = "参加者" & Chr(10) & "サイン"
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("C14:K14")
        .Merge
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("L14")
        .Value = "作業員"
        .Font.Bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("M14:N14")
        .Merge
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("O14")
        .Value = "名"
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .BorderAround xlContinuous, xlMedium
    End With

    With ws.Range("P14")
        .BorderAround xlContinuous, xlMedium
    End With
End Sub

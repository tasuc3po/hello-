Attribute VB_Name = "KY表マクロ"
Option Explicit

'==============================================
' KY表 自動作成マクロ  Ver 3.0
' 公民連携沖縄株式会社
' 実際のExcelファイル（37行）構造に完全準拠
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
           "「" & SHEET_INPUT & "」シートに記入後、" & vbCrLf & _
           "「KY表を作成する」ボタンを押してください。", _
           vbInformation, "KY表マクロ Ver 3.0"
End Sub

'----------------------------------------------
' KY表にデータを転記して印刷プレビューを開く
'----------------------------------------------
Public Sub KY表作成()
    Dim wsIn As Worksheet
    Dim wsKY As Worksheet

    On Error Resume Next
    Set wsIn = ThisWorkbook.Sheets(SHEET_INPUT)
    Set wsKY = ThisWorkbook.Sheets(SHEET_KY)
    On Error GoTo 0

    If wsIn Is Nothing Or wsKY Is Nothing Then
        MsgBox "先に「セットアップ」を実行してください。", vbExclamation
        Exit Sub
    End If

    Dim m As String, d As String
    m = Trim(CStr(wsIn.Range("C3").Value))
    d = Trim(CStr(wsIn.Range("C4").Value))
    If m = "" Then MsgBox "「月」を入力してください。", vbExclamation: Exit Sub
    If d = "" Then MsgBox "「日」を入力してください。", vbExclamation: Exit Sub

    Application.ScreenUpdating = False

    Dim reiwaYear As Integer
    reiwaYear = Year(Now) - REIWA_BASE
    Dim yobi As String
    yobi = GetYobi(CInt(m), CInt(d))

    ' ① H1:L4  日付（令和X年 M月 D日（Y曜日））
    wsKY.Range("H1").Value = "令和 " & reiwaYear & " 年　" & _
                              CInt(m) & " 月　" & CInt(d) & " 日（" & yobi & " 曜日）"

    ' ② M1:N4  天候
    wsKY.Range("M1").Value = "天候　" & wsIn.Range("C5").Value

    ' ③ E5:N7  グループの作業内容
    wsKY.Range("E5").Value = wsIn.Range("C6").Value

    ' ④ A11:G13  危険のポイント①
    wsKY.Range("A11").Value = wsIn.Range("C7").Value

    ' ⑤ H11:N13  私達はこうする①
    wsKY.Range("H11").Value = wsIn.Range("C8").Value

    ' ⑥ A14:G16  危険のポイント②
    wsKY.Range("A14").Value = wsIn.Range("C9").Value

    ' ⑦ H14:N16  私達はこうする②
    wsKY.Range("H14").Value = wsIn.Range("C10").Value

    ' ⑧ E23:N26  本日の安全目標
    wsKY.Range("E23").Value = wsIn.Range("C11").Value

    ' ⑨ J32:N34  リーダー名
    Dim leader As String
    leader = Trim(CStr(wsIn.Range("C12").Value))
    If leader <> "" Then
        wsKY.Range("J32").Value = "リーダー　" & leader
    End If

    ' ⑩ C35:K37  参加者氏名
    wsKY.Range("C35").Value = wsIn.Range("C13").Value

    ' ⑪ L35:M37  作業員数
    Dim cnt As String
    cnt = Trim(CStr(wsIn.Range("C14").Value))
    If cnt <> "" Then
        wsKY.Range("L35").Value = "作業員　" & cnt
    End If

    Application.ScreenUpdating = True
    wsKY.Activate
    wsKY.PrintPreview
    wsIn.Activate
End Sub

'----------------------------------------------
' 曜日取得
'----------------------------------------------
Private Function GetYobi(m As Integer, d As Integer) As String
    Dim arr As Variant
    arr = Array("日", "月", "火", "水", "木", "金", "土")
    On Error GoTo ErrExit
    GetYobi = arr(Weekday(DateSerial(Year(Now), m, d)) - 1)
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

    ' タイトル行
    With ws.Range("A1:E1")
        .Merge
        .Value = "KY表　入力フォーム　Ver 3.0"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Rows(1).RowHeight = 32

    With ws.Range("A2:E2")
        .Merge
        .Value = "※ 黄色セルに入力後、「KY表を作成する」ボタンを押してください"
        .Font.Size = 9
        .Font.Color = RGB(150, 150, 150)
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Rows(2).RowHeight = 16

    ' 入力項目
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
    labels(10) = "参加者氏名"
    labels(11) = "作業員数（名）"

    heights(0)  = 22:  heights(1)  = 22:  heights(2)  = 22
    heights(3)  = 50:  heights(4)  = 60:  heights(5)  = 60
    heights(6)  = 60:  heights(7)  = 60:  heights(8)  = 40
    heights(9)  = 22:  heights(10) = 60:  heights(11) = 22

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

    ' 天候デフォルト値
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

    ' 作成ボタン
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("C16").Left + 5, ws.Range("C16").Top + 8, 200, 38)
    With shp
        .Name = "btnKY"
        .TextFrame.Characters.Text = "▶  KY表を作成する"
        .TextFrame.Characters.Font.Size = 13
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
' 実際のExcelファイル（37行×A〜N列）に完全準拠
'==============================================================
Private Sub CreateKYSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_KY

    ' ---- ページ設定（A4 横）----
    With ws.PageSetup
        .Orientation   = xlLandscape
        .PaperSize     = xlPaperA4
        .Zoom          = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin     = Application.InchesToPoints(0.748)
        .BottomMargin  = Application.InchesToPoints(0.748)
        .LeftMargin    = Application.InchesToPoints(0)
        .RightMargin   = Application.InchesToPoints(0)
        .CenterHorizontally = True
    End With

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' ---- 列幅（実測値） ----
    ws.Columns("A").ColumnWidth = 6.5
    ws.Columns("B").ColumnWidth = 9.0
    ws.Columns("C").ColumnWidth = 8.11
    ws.Columns("D").ColumnWidth = 8.11
    ws.Columns("E").ColumnWidth = 8.11
    ws.Columns("F").ColumnWidth = 8.11
    ws.Columns("G").ColumnWidth = 10.5
    ws.Columns("H").ColumnWidth = 9.0
    ws.Columns("I").ColumnWidth = 8.11
    ws.Columns("J").ColumnWidth = 8.11
    ws.Columns("K").ColumnWidth = 8.11
    ws.Columns("L").ColumnWidth = 27.125
    ws.Columns("M").ColumnWidth = 10.625
    ws.Columns("N").ColumnWidth = 6.5

    ' ---- 行高（実測値）----
    Dim rh As Integer
    For rh = 1 To 37
        ws.Rows(rh).RowHeight = 13.5
    Next rh
    ws.Rows(5).RowHeight  = 14.25
    ws.Rows(7).RowHeight  = 23.25
    ws.Rows(10).RowHeight = 14.25
    ws.Rows(13).RowHeight = 14.25
    ws.Rows(16).RowHeight = 14.25
    ws.Rows(19).RowHeight = 14.25
    ws.Rows(22).RowHeight = 14.25
    ws.Rows(23).RowHeight = 14.25
    ws.Rows(26).RowHeight = 14.25
    ws.Rows(30).RowHeight = 12.75
    ws.Rows(31).RowHeight = 12.75
    ws.Rows(34).RowHeight = 14.25
    ws.Rows(37).RowHeight = 14.25

    ' ============================================================
    ' 行1〜4：タイトル・日付・天候
    ' ============================================================

    ' A1:G4  危険予知活動表
    ws.Range("A1:G4").Merge
    With ws.Range("A1")
        .Value = "危険予知活動表"
        .Font.Name = "HGP創英角ｺﾞｼｯｸUB"
        .Font.Size = 48
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("A1:G4").BorderAround xlContinuous, xlMedium

    ' H1:L4  日付（転記時に上書き）
    ws.Range("H1:L4").Merge
    Dim reiwaYear As Integer
    reiwaYear = Year(Now) - REIWA_BASE
    With ws.Range("H1")
        .Value = "令和 " & reiwaYear & " 年　　月　　日（　　曜日）"
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 26
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("H1:L4").BorderAround xlContinuous, xlMedium

    ' M1:N4  天候
    ws.Range("M1:N4").Merge
    With ws.Range("M1")
        .Value = "天候"
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 18
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("M1:N4").BorderAround xlContinuous, xlMedium

    ' ============================================================
    ' 行5〜7：グループの作業内容
    ' ============================================================

    ' A5:D7  ラベル
    ws.Range("A5:D7").Merge
    With ws.Range("A5")
        .Value = "グループの作業内容"
        .Font.Name = "HGPｺﾞｼｯｸE"
        .Font.Size = 20
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
    End With
    ws.Range("A5:D7").BorderAround xlContinuous, xlMedium

    ' E5:N7  入力エリア
    ws.Range("E5:N7").Merge
    With ws.Range("E5")
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 16
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
    End With
    ws.Range("E5:N7").BorderAround xlContinuous, xlMedium

    ' ============================================================
    ' 行8〜10：KYヘッダー
    ' ============================================================

    ' A8:G10  危険のポイント（見出し）
    ws.Range("A8:G10").Merge
    With ws.Range("A8")
        .Value = "危険のポイント"
        .Font.Name = "HGPｺﾞｼｯｸE"
        .Font.Size = 28
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("A8:G10").BorderAround xlContinuous, xlMedium

    ' H8:N10  私達はこうする（見出し）
    ws.Range("H8:N10").Merge
    With ws.Range("H8")
        .Value = "私達はこうする"
        .Font.Name = "HGPｺﾞｼｯｸE"
        .Font.Size = 28
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("H8:N10").BorderAround xlContinuous, xlMedium

    ' ============================================================
    ' 行11〜22：危険ポイント入力エリア（4ペア、①②のみ転記）
    ' ============================================================
    Dim pairRanges(3, 1) As String
    pairRanges(0, 0) = "A11:G13" : pairRanges(0, 1) = "H11:N13"
    pairRanges(1, 0) = "A14:G16" : pairRanges(1, 1) = "H14:N16"
    pairRanges(2, 0) = "A17:G19" : pairRanges(2, 1) = "H17:N19"
    pairRanges(3, 0) = "A20:G22" : pairRanges(3, 1) = "H20:N22"

    Dim p As Integer
    For p = 0 To 3
        Dim leftCell  As String
        Dim rightCell As String
        leftCell  = Split(pairRanges(p, 0), ":")(0)
        rightCell = Split(pairRanges(p, 1), ":")(0)

        ws.Range(pairRanges(p, 0)).Merge
        With ws.Range(leftCell)
            .Font.Name = "ＭＳ Ｐゴシック"
            .Font.Size = 14
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
        End With
        ws.Range(pairRanges(p, 0)).BorderAround xlContinuous, xlMedium

        ws.Range(pairRanges(p, 1)).Merge
        With ws.Range(rightCell)
            .Font.Name = "ＭＳ Ｐゴシック"
            .Font.Size = 14
            .VerticalAlignment = xlVAlignTop
            .WrapText = True
        End With
        ws.Range(pairRanges(p, 1)).BorderAround xlContinuous, xlMedium
    Next p

    ' ============================================================
    ' 行23〜26：本日の安全目標
    ' ============================================================

    ' A23:D26  ラベル
    ws.Range("A23:D26").Merge
    With ws.Range("A23")
        .Value = "本日の安全目標"
        .Font.Name = "HGS創英角ｺﾞｼｯｸUB"
        .Font.Size = 20
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
    End With
    ws.Range("A23:D26").BorderAround xlContinuous, xlMedium

    ' E23:N26  入力エリア
    ws.Range("E23:N26").Merge
    With ws.Range("E23")
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 16
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
    End With
    ws.Range("E23:N26").BorderAround xlContinuous, xlMedium

    ' ============================================================
    ' 行27〜31：チェックリスト
    ' ============================================================
    ws.Range("A27:N31").Merge
    With ws.Range("A27")
        .Value = "□ 現地の安全確認は実施したか　　□ 体調不良者はいないか （無し・対処済）　　□ 服装・装備は万全か" & Chr(10) & _
                 "□ 天候や視界に問題はないか （無し・対処済）　　□ 事前の予定と異なった要素はないか （無し・対処済）"
        .Font.Name = "HGPｺﾞｼｯｸE"
        .Font.Size = 14
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
    End With
    ws.Range("A27:N31").BorderAround xlContinuous, xlMedium

    ' ============================================================
    ' 行32〜34：会社名・リーダー
    ' ============================================================

    ' A32:B34  「会社名」ラベル
    ws.Range("A32:B34").Merge
    With ws.Range("A32")
        .Value = "会社名"
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 20
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("A32:B34").BorderAround xlContinuous, xlMedium

    ' C32:I34  会社名
    ws.Range("C32:I34").Merge
    With ws.Range("C32")
        .Value = COMPANY
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 22
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("C32:I34").BorderAround xlContinuous, xlMedium

    ' J32:N34  リーダー（転記時に上書き）
    ws.Range("J32:N34").Merge
    With ws.Range("J32")
        .Value = "リーダー"
        .Font.Name = "HGPｺﾞｼｯｸE"
        .Font.Size = 16
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    ws.Range("J32:N34").BorderAround xlContinuous, xlMedium

    ' ============================================================
    ' 行35〜37：参加者・作業員数
    ' ============================================================

    ' A35:B37  「参加者 サイン」ラベル
    ws.Range("A35:B37").Merge
    With ws.Range("A35")
        .Value = "参加者" & Chr(10) & "サイン"
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 18
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
    End With
    ws.Range("A35:B37").BorderAround xlContinuous, xlMedium

    ' C35:K37  参加者氏名入力エリア
    ws.Range("C35:K37").Merge
    With ws.Range("C35")
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 14
        .VerticalAlignment = xlVAlignTop
        .WrapText = True
    End With
    ws.Range("C35:K37").BorderAround xlContinuous, xlMedium

    ' L35:M37  「作業員」ラベル（転記時に人数を付加）
    ws.Range("L35:M37").Merge
    With ws.Range("L35")
        .Value = "作業員"
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 18
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("L35:M37").BorderAround xlContinuous, xlMedium

    ' N35:N37  「名」
    ws.Range("N35:N37").Merge
    With ws.Range("N35")
        .Value = "名"
        .Font.Name = "HGPｺﾞｼｯｸE"
        .Font.Size = 26
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Range("N35:N37").BorderAround xlContinuous, xlMedium

End Sub

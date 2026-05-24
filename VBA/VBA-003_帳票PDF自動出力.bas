' [VBA-003] 帳票PDF自動出力マクロ
' 用途: 指定範囲をPDFとして保存（請求書・納品書等）
' カスタマイズ: RANGE_ADDRESS, SAVE_FOLDER を変更

Option Explicit

Const RANGE_ADDRESS As String = "A1:H40"               ' ← 印刷範囲
Const SAVE_FOLDER   As String = "C:\work\pdf_output\"  ' ← 保存先

Sub ExportToPDF()
    Dim ws       As Worksheet
    Dim fileName As String
    Dim savePath As String

    Set ws = ActiveSheet

    ' ファイル名：シート名 + 今日の日付
    fileName = ws.Name & "_" & Format(Date, "YYYYMMDD") & ".pdf"
    savePath = SAVE_FOLDER & fileName

    ws.Range(RANGE_ADDRESS).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False

    MsgBox "PDF出力完了: " & savePath, vbInformation
End Sub

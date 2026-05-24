' [VBA-001] Excel集計マクロ（基本型）
' 用途: 指定シートのデータを集計して別シートに出力
' カスタマイズ: INPUT_SHEET, OUTPUT_SHEET, COL_TARGET を変更

Option Explicit

Const INPUT_SHEET  As String = "データ"
Const OUTPUT_SHEET As String = "集計"
Const COL_TARGET   As Long = 3   ' 集計対象列番号

Sub CollectData()
    Dim wsIn  As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object

    Set wsIn  = ThisWorkbook.Sheets(INPUT_SHEET)
    Set wsOut = ThisWorkbook.Sheets(OUTPUT_SHEET)
    Set dict  = CreateObject("Scripting.Dictionary")

    lastRow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row

    ' データ読み込み
    For i = 2 To lastRow
        Dim key As String
        key = CStr(wsIn.Cells(i, COL_TARGET).Value)
        If dict.exists(key) Then
            dict(key) = dict(key) + wsIn.Cells(i, COL_TARGET + 1).Value
        Else
            dict.Add key, wsIn.Cells(i, COL_TARGET + 1).Value
        End If
    Next i

    ' 出力
    wsOut.Cells.ClearContents
    wsOut.Cells(1, 1).Value = "項目"
    wsOut.Cells(1, 2).Value = "合計"

    Dim row As Long
    row = 2
    Dim k As Variant
    For Each k In dict.Keys
        wsOut.Cells(row, 1).Value = k
        wsOut.Cells(row, 2).Value = dict(k)
        row = row + 1
    Next k

    MsgBox "集計完了！" & (row - 2) & "件", vbInformation
End Sub

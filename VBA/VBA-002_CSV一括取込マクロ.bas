' [VBA-002] CSV一括取込マクロ
' 用途: フォルダ内の全CSVを読み込んで1シートに結合
' カスタマイズ: FOLDER_PATH, DELIMITER を変更

Option Explicit

Const FOLDER_PATH As String = "C:\work\csv_files\"   ' ← 変更
Const DELIMITER   As String = ","

Sub ImportAllCSV()
    Dim wsOut    As Worksheet
    Dim filePath As String
    Dim fileNum  As Integer
    Dim lineData As String
    Dim cols()   As String
    Dim outRow   As Long
    Dim isFirst  As Boolean

    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "CSV結合_" & Format(Now, "MMDD_HHmm")
    outRow = 1
    isFirst = True

    filePath = Dir(FOLDER_PATH & "*.csv")
    Do While filePath <> ""
        fileNum = FreeFile
        Open FOLDER_PATH & filePath For Input As #fileNum

        Dim skipHeader As Boolean
        skipHeader = Not isFirst   ' 2ファイル目以降はヘッダーをスキップ

        Do While Not EOF(fileNum)
            Line Input #fileNum, lineData
            If skipHeader Then
                skipHeader = False
            Else
                cols = Split(lineData, DELIMITER)
                Dim c As Long
                For c = 0 To UBound(cols)
                    wsOut.Cells(outRow, c + 1).Value = cols(c)
                Next c
                outRow = outRow + 1
            End If
        Loop

        Close #fileNum
        isFirst = False
        filePath = Dir
    Loop

    MsgBox "取込完了！" & (outRow - 1) & "行", vbInformation
End Sub

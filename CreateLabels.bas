Attribute VB_Name = "Module1"
Sub CreateLabels()
    Dim sourceSheet As Worksheet
    Dim templateSheet As Worksheet
    Dim newSheet As Worksheet
    
    Dim dataRow As Long
    Dim formCount As Long
    Dim sheetIndex As Long
    Dim lastRow As Long
    
    Dim rowOffsets As Variant
    Dim colOffsets As Variant
    Dim r As Integer, c As Integer
    Dim topLeftCell As Range
    Dim newSheetName As String

    ' Define sheets
    Set sourceSheet = ThisWorkbook.Sheets("data")      ' source data
    Set templateSheet = ThisWorkbook.Sheets("form")    ' label template

    ' Find last row in column B (Full Name column)
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 2).End(xlUp).Row
    
    ' Initialize counters
    dataRow = 2
    formCount = 0
    sheetIndex = 1
    
    ' Positions of label blocks (adjust to match your template layout)
    rowOffsets = Array(9, 33, 57)   ' 3 rows of labels per page
    colOffsets = Array("D", "M")    ' 2 columns per row
    
    ' Main loop: process all winners until lastRow
    Do While dataRow <= lastRow And sourceSheet.Cells(dataRow, 2).Value <> ""
        
        ' Copy template sheet
        templateSheet.Copy After:=Sheets(Sheets.Count)
        Set newSheet = ActiveSheet
        
        ' Assign safe sheet name
        newSheetName = "Form Page " & sheetIndex
        On Error Resume Next
        newSheet.Name = newSheetName
        On Error GoTo 0
        
        ' Fill up to 6 labels per page (2×3 grid)
        For r = 0 To 2
            For c = 0 To 1
                If dataRow > lastRow Then Exit For   ' stop if no more data
                
                ' Top-left cell of this label block
                Set topLeftCell = newSheet.Range(colOffsets(c) & rowOffsets(r))
                
                ' Map fields from "data" into the label
                topLeftCell.Value = sourceSheet.Cells(dataRow, 2).Value   ' Full Name (col B)
                topLeftCell.Offset(1, 0).Value = sourceSheet.Cells(dataRow, 3).Value ' Street & Number (col C)
                topLeftCell.Offset(2, 0).Value = sourceSheet.Cells(dataRow, 5).Value ' Floor (col E)
                topLeftCell.Offset(3, 0).Value = sourceSheet.Cells(dataRow, 7).Value ' City & Region (col G)
                topLeftCell.Offset(4, 0).Value = sourceSheet.Cells(dataRow, 6).Value ' Postal Code (col F)
                topLeftCell.Offset(5, 0).Value = sourceSheet.Cells(dataRow, 4).Value ' Contact Phone (col D)
                
                ' Next winner
                dataRow = dataRow + 1
                formCount = formCount + 1
            Next c
            If dataRow > lastRow Then Exit For
        Next r
        
        sheetIndex = sheetIndex + 1
    Loop
    
    ' Report result
    MsgBox formCount & " labels created across " & (sheetIndex - 1) & " sheets.", vbInformation

End Sub


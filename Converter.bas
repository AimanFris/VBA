Sub convert()
Paste Values Only With PasteSpecial

    Dim ExistingTable As Range
    Dim OutputRange As Range
    Dim OutRow As Long
    Dim r As Long, c As Long
    Dim test As Range

'Copy A Range of Data
  Set ExistingTable = Application.InputBox(Prompt:="Select Cell that contain data", Type:=8).CurrentRegion
  ActiveWorkbook.Sheets.Add.Name = "Table2"
 
  
  ExistingTable.copy Worksheets("Table2").Range("A1")
  Worksheets("Table2").Range("A1").CurrentRegion.Value = ExistingTable.Value
 

'Clear Clipboard (removes "marching ants" around your original data set)
  Application.CutCopyMode = False
  

    On Error Resume Next
'    Prompt User to select cell
    Set ExistingTable = Worksheets("Table2").Range("A1").CurrentRegion

    If ExistingTable.Count = 1 Or ExistingTable.Rows.Count < 3 Then
        MsgBox "Select a cell within the summary table.", vbCritical
        Exit Sub
    End If
    
    ExistingTable.Select

    
'   DELETE EMPTY ROW
    Row = 3
    For r = 3 To ExistingTable.Rows.Count * 2
        
        If IsEmpty(ExistingTable.Cells(Row, 2)) Then
        ExistingTable.Cells(Row, 2).EntireRow.Delete
        Row = Row - 1
        End If
        Row = Row + 1
    Next r
    
    
'   REPLACE EMPTY CELL WITH ZERO
    For r = 3 To ExistingTable.Rows.Count
        For c = 3 To ExistingTable.Columns.Count
        
        If IsEmpty(ExistingTable.Cells(r, c)) Then
        ExistingTable.Cells(r, c) = 0
        End If
        
        Next c
    Next r
    ExistingTable.Select
    
    
    Set OutputRange = Application.InputBox(Prompt:="Select an empty cell for the output. ", Type:=8)
'   Convert the range
    OutRow = 2
    col_p = 3
    col_a = 4
    Application.ScreenUpdating = False
    OutputRange.Range("A1:E5") = Array("Date", "Line", "Style", "Plan", "Actual")
    For r = 3 To ExistingTable.Rows.Count
        col_p = 3
        col_a = 4
        For c = 3 To (ExistingTable.Columns.Count / 2) + 1
'            Date
            OutputRange.Cells(OutRow, 1) = Format(ExistingTable.Cells(1, col_p), "mm/dd/yyyy")

'            Line
            OutputRange.Cells(OutRow, 2) = "Line" & ExistingTable.Cells(r, 1)
'            Style
            OutputRange.Cells(OutRow, 3) = ExistingTable.Cells(r, 2)
            
'           Plan
            OutputRange.Cells(OutRow, 4) = ExistingTable.Cells(r, col_p)
            OutputRange.Cells(OutRow, 4).NumberFormat = ExistingTable.Cells(r, col_p).NumberFormat
            
'           Actual
            OutputRange.Cells(OutRow, 5) = ExistingTable.Cells(r, col_a)
            OutputRange.Cells(OutRow, 5).NumberFormat = ExistingTable.Cells(r, col_a).NumberFormat
            
            OutRow = OutRow + 1
            col_p = col_p + 2
            col_a = col_a + 2
        Next c
    Next r

End Sub



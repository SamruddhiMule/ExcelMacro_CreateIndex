Sub CreateModifiedIndexSheet()
    Dim ws As Worksheet
    Dim indexSheet As Worksheet
    Dim lastRow As Long
    Dim srNo As Long
    
    ' Create a new worksheet named "Index" or use an existing one
    On Error Resume Next
    Set indexSheet = Sheets("Index")
    On Error GoTo 0
    
    If indexSheet Is Nothing Then
        Set indexSheet = Sheets.Add(Before:=Sheets(1)) ' Insert at the beginning
        indexSheet.Name = "Index"
    End If
    
    ' Add title in one cell with larger font size, light yellow background, and center alignment
    With indexSheet.Range("A1:C1")
        .Merge
        .Value = "Table of Contents"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(255, 255, 153) ' Light yellow background
        .HorizontalAlignment = xlCenter ' Center-aligned
    End With
    
    ' Delete empty row after the table title
    If IsEmpty(indexSheet.Cells(2, 1).Value) Then
        indexSheet.Rows(2).Delete Shift:=xlUp
    End If
    
    ' Add column headers with light yellow background
    indexSheet.Range("A3:C3").Interior.Color = RGB(255, 255, 153) ' Light yellow background
    indexSheet.Range("A3:C3").HorizontalAlignment = xlCenter ' Center-aligned
    
    indexSheet.Cells(3, 1).Value = "Sr No."
    indexSheet.Cells(3, 1).Font.Bold = True
    
    indexSheet.Cells(3, 2).Value = "Particulars"
    indexSheet.Cells(3, 2).Font.Bold = True
    
    indexSheet.Cells(3, 3).Value = "Description"
    indexSheet.Cells(3, 3).Font.Bold = True
    
    ' Loop through all worksheets in the workbook
    srNo = 1
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in the index sheet
        lastRow = indexSheet.Cells(indexSheet.Rows.Count, 1).End(xlUp).Row
        
        ' Add data to the index sheet
        With indexSheet.Cells(lastRow + 1, 1)
            .Value = srNo
            .Hyperlinks.Add _
                Anchor:=.Offset(0, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
        End With
        
        indexSheet.Cells(lastRow + 1, 2).Value = ws.Name
        indexSheet.Cells(lastRow + 1, 3).Value = ws.Range("A1").Value
        
        srNo = srNo + 1
    Next ws
    
    ' Apply light yellow background color to the entire index table
    indexSheet.Range("A3:C" & indexSheet.Cells(indexSheet.Rows.Count, 1).End(xlUp).Row).Interior.Color = RGB(255, 255, 153)
    indexSheet.Range("A3:C" & indexSheet.Cells(indexSheet.Rows.Count, 1).End(xlUp).Row).HorizontalAlignment = xlCenter ' Center-aligned
End Sub







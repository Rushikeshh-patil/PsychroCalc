Sub ExportTable4a()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblCount As Integer
    Dim newFolderPath As String
    Dim i As Integer
    Dim tbl2 As Range
    Dim ws1 As Worksheet
    Dim table4aname As String
    Dim sheetCount As Integer
    Dim sheetNum As Integer
    Dim progressPct As Integer
    Dim progressBar As Object
    
    ' Get the path of the current workbook
    newFolderPath = ThisWorkbook.Path & "/Table 4a Schedules/"
    
    On Error Resume Next
    Kill newFolderPath & "*.*"
    RmDir newFolderPath
    On Error GoTo 0
    
    ' Create folder
    MkDir newFolderPath
    
    ' Get the number of worksheets in the current workbook whose name contains "AHU"
    sheetCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "AHU", vbTextCompare) > 0 Then
            sheetCount = sheetCount + 1
        End If
    Next ws
    
    ' Create progress bar
    Set progressBar = CreateObject("MSComctlLib.ProgressBar")
    progressBar.Min = 0
    progressBar.Max = sheetCount
    progressBar.Value = 0
    progressBar.Visible = True
    
    ' Loop through each worksheet in the current workbook whose name contains "AHU"
    sheetNum = 0
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "AHU", vbTextCompare) > 0 Then
            ' Loop through each table in the current worksheet
            tblCount = ws.ListObjects.Count
            If tblCount > 0 Then
                For i = 1 To tblCount
                    Set tbl = ws.ListObjects(i)
                    ' Check if the table name contains "table4a"
                    If InStr(1, tbl.Name, "table4a", vbTextCompare) > 0 Then
                        ' Create a new workbook and copy the table to it
                        Workbooks.Add
                        tbl.Range.Columns("A:T").Copy
                        ActiveSheet.Range("A1").PasteSpecial xlPasteValues
                        Set ws1 = ActiveSheet
                        Set tbl2 = Range("A1").CurrentRegion
                        table4aname = ws.Name & "Table4A"
                        ws1.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl2).Name = table4aname
                        With ActiveSheet.ListObjects(1)
                            ' Center all cell values in the table
                            .Range.HorizontalAlignment = xlCenter
                            ' Apply Arial font with size 12 to the entire table
                            .Range.Font.Name = "Arial"
                            .Range.Font.Size = 12
                            ' Automatically adjust the width of all columns to fit their content
                            .Range.Columns.AutoFit
                        End With
                        
                        ' Save the new workbook in the "Tables" folder
                        ActiveWorkbook.SaveAs newFolderPath & ws.Name & "_Table4a.xls"
                        
                        ' Close the new workbook
                        ActiveWorkbook.Close False
                    End If
                Next i
            End If
            sheetNum = sheetNum + 1
            progressPct = (sheetNum / sheetCount) * 100
            progressBar.Value = progressPct
        End If
    Next ws
    
    ' Close progress bar
    progressBar.Visible = False
End Sub

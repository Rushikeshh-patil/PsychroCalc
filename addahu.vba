Sub ahuadd()

    Dim lastRow As Long
    lastRow = Sheets("Psych").Range("Table7[TAG]").Rows.Count
    Worksheets("Generic").Visible = xlSheetVisible


    ' Copy the Generic worksheet and rename it with the next AHU number
    Sheets("Generic").Copy Before:=Sheets(1)
    Dim newAhuName As String
    newAhuName = "AHU " & lastRow + 1
    ActiveSheet.Name = newAhuName
    Set refsheet = ThisWorkbook.Sheets(newAhuName)
    Set scfm = refsheet.Range("AW2")
    Set rcfm = refsheet.Range("AX2")
    Set ocfm = refsheet.Range("AY2")
    Set gpm = refsheet.Range("BA2")
    Set mbh = refsheet.Range("BB2")

    ' Add a row to the table
    With Sheets("Psych").Range("Table7[TAG]").ListObject
        Dim newRow As ListRow
        Set newRow = .ListRows.Add
        newRow.Range(1, 1).Value = newAhuName
        newRow.Range(1, 2).Formula = "='" & refsheet.Name & "'!" & scfm.Address(ReferenceStyle:=xlR1C1)
        newRow.Range(1, 4).Formula = "='" & refsheet.Name & "'!" & rcfm.Address(ReferenceStyle:=xlR1C1)
        newRow.Range(1, 6).Value = .ListRows(.ListRows.Count - 1).Range(1, 6).Value
        newRow.Range(1, 7).Value = .ListRows(.ListRows.Count - 1).Range(1, 7).Value
        newRow.Range(1, 8).FormulaR1C1 = "='" & refsheet.Name & "'!" & ocfm.Address(ReferenceStyle:=xlR1C1)
        newRow.Range(1, 19).Value = .ListRows(.ListRows.Count - 1).Range(1, 19).Value
        newRow.Range(1, 20).Value = .ListRows(.ListRows.Count - 1).Range(1, 20).Value
        newRow.Range(1, 55).Value = .ListRows(.ListRows.Count - 1).Range(1, 55).Value
        newRow.Range(1, 56).Value = .ListRows(.ListRows.Count - 1).Range(1, 56).Value
        newRow.Range(1, 57).Formula = "='" & refsheet.Name & "'!" & gpm.Address(ReferenceStyle:=xlR1C1)
        newRow.Range(1, 58).Formula = "='" & refsheet.Name & "'!" & mbh.Address(ReferenceStyle:=xlR1C1)
    End With
    
    With Sheets("INPUT_OUTPUTS").Range("AHU_Options").ListObject
        Dim newRow As ListRow
        Set newRow = .ListRows.Add
        newRow.Range(1, 1).Value = newAhuName
    End with
        
    
    'hide the "Generic" worksheet
    Worksheets("Generic").Visible = xlSheetHidden
    ThisWorkbook.Sheets("Psych").Activate
    MsgBox "Added New AHU!"
End Sub


Sub removeAHU()
    'get the last row in the "Table7[TAG]" table
    Dim lastRow As Long
    lastRow = Sheets("Psych").Range("Table7[TAG]").Rows.Count
    'get the name of the last added AHU sheet
    Dim lastSheetName As String
    lastSheetName = Sheets("Psych").Range("Table7[TAG]").Cells(lastRow, 1).Value

    'prompt user to confirm deletion
    Dim confirm As Integer
    confirm = MsgBox("Do you really want to delete " & lastSheetName & "?", vbYesNo + vbQuestion, "Confirm Deletion")
    If confirm = vbYes Then
        'delete the last row in the "Table7[TAG]" table
        Sheets("Psych").Range("Table7[TAG]").ListObject.ListRows(lastRow).Delete
        'delete the last added AHU sheet
        Application.DisplayAlerts = False
        Sheets(lastSheetName).Delete
        Application.DisplayAlerts = True
    Else
        'if user chooses not to delete, exit the function
        Exit Sub
    End If
End Sub

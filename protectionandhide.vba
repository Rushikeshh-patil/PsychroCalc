Sub Protection(wSht As Worksheet, ProtAction As Boolean)
    'This sub is called to Protect and Unprotect worksheets
   
    If ProtAction = True Then
        wSht.Protect Password:=strPassword  'Protect the worksheet
    Else
        wSht.Unprotect Password:=strPassword    'Unprotect the worksheet
    End If
End Sub

Sub GroupAndHidepreheat()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Psych")
    Call Protection(ws, False)

    'Define the range of columns to be grouped
    Set Rng = ws.Range("At:az")

    'Check the current hidden state of the columns
    If Rng.EntireColumn.Hidden = True Then
        'Ungroup the columns and show them
        Rng.EntireColumn.Ungroup
        Rng.EntireColumn.Hidden = False
    Else
        'Group the columns and hide them
        Rng.EntireColumn.Group
        Rng.EntireColumn.Hidden = True
    End If
    Call Protection(ws, True)
End Sub

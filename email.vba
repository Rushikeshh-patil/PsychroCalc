Sub create_paragraph()
    Dim num_rows As Integer
    Dim inputs_outputs As Worksheet
    Dim table7 As ListObject
    Dim paragraph As String
    Dim row As Range
    
    Set inputs_outputs = ThisWorkbook.Sheets("input_outputs")
    Set table7 = ThisWorkbook.Sheets("Psych").ListObjects("table7")
    Set ahuOptions = ThisWorkbook.Sheets("INPUT_OUTPUTS").ListObjects("AHU_Options")
    num_rows = table7.Range.Rows.Count

    paragraph = "Hi" & vbNewLine & vbNewLine & "I want to get selections for " & num_rows - 2 & " AHUs." & vbNewLine & vbNewLine & "The outside air conditions are as follows:" & vbNewLine & "Summer DB: " & inputs_outputs.Range("C8").Value & vbNewLine & "Summer WB: " & inputs_outputs.Range("C9").Value & vbNewLine & "Winter DB: " & inputs_outputs.Range("F8").Value & vbNewLine & vbNewLine & "The water side temperatures are as follows:" & vbNewLine & "CHWS: " & inputs_outputs.Range("C14").Value & vbNewLine & "CHWR: " & inputs_outputs.Range("C15").Value & vbNewLine & "HHWS: " & inputs_outputs.Range("C16").Value & vbNewLine & "HHWR: " & inputs_outputs.Range("C17").Value & vbNewLine & vbNewLine & "The AHU information is as follows:" & vbNewLine
    
    For Each row In table7.DataBodyRange.Rows
        paragraph = paragraph & "     - " & row.Cells(1).Value & vbNewLine & "          - Supply Air CFM: " & row.Cells(2).Value & vbNewLine & "          - Return Air CFM: " & row.Cells(4).Value & vbNewLine & "          - OA CFM: " & row.Cells(8).Value & vbNewLine & "          - Cooling LAT: " & row.Cells(19).Value & " DB and " & row.Cells(20).Value & " WB" & vbNewLine & "          - Room set point: " & row.Cells(6).Value & " DB and " & row.Cells(7).Value & " WB" & vbNewLine
    Next row
    
    For Each row In ahuOptions.DataBodyRange.Rows
        paragraph = paragraph & "     - " & row.Cells(1).Value & vbNewLine & "          - Discharge Cinfiguration: " & row.Cells(1).Value & vbNewLine
    Next row

    Dim FilePath As String
    FilePath = ThisWorkbook.Path & "\SelectionRequest.txt"
    
    If Dir(FilePath) <> "" Then
        Kill FilePath
    End If
    
    Open FilePath For Output As #1
    Print #1, paragraph
    Close #1
    
    MsgBox "Done! Check folder"

End Sub


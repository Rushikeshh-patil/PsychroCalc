Sub create_paragraph()
    Dim num_rows As Integer
    Dim inputs_outputs As Worksheet
    Dim table7 As ListObject
    Dim paragraph As String
    Dim row As Range
    
    Set inputs_outputs = ThisWorkbook.Sheets("input_outputs")
    Set table7 = ThisWorkbook.Sheets("Psych").ListObjects("table7")
    num_rows = table7.Range.Rows.Count
    paragraph = "Hi" & vbNewLine & vbNewLine & "I want to get selections for " & num_rows & " AHUs." & vbNewLine & vbNewLine & "The outside air conditions are as follows:" & vbNewLine & "Summer DB: " & inputs_outputs.Range("C5").Value & vbNewLine & "Summer WB: " & inputs_outputs.Range("C6").Value & vbNewLine & "Winter DB: " & inputs_outputs.Range("F5").Value & vbNewLine & vbNewLine & "The water side temperatures are as follows:" & vbNewLine & "CHWS: " & inputs_outputs.Range("C11").Value & vbNewLine & "CHWR: " & inputs_outputs.Range("C12").Value & vbNewLine & "HHWS: " & inputs_outputs.Range("C13").Value & vbNewLine & "HHWR: " & inputs_outputs.Range("C14").Value & vbNewLine & vbNewLine & "The AHU information is as follows:" & vbNewLine
    For Each row In table7.DataBodyRange.Rows
        paragraph = paragraph & "     - " & row.Cells(1).Value & vbNewLine & "          - Supply Air CFM: " & row.Cells(2).Value & vbNewLine & "          - Return Air CFM: " & row.Cells(4).Value & vbNewLine & "          - OA CFM: " & row.Cells(8).Value & vbNewLine & "          - LAT: " & row.Cells(19).Value & " DB and " & row.Cells(20).Value & " WB" & vbNewLine & "          - Room set point: " & row.Cells(6).Value & " DB and " & row.Cells(7).Value & " WB" & vbNewLine
    Next row
    Dim FilePath As String
    FilePath = ThisWorkbook.Path & "\AHU_information.txt"
    Open FilePath For Output As #1
    Print #1, paragraph
    Close #1

End Sub


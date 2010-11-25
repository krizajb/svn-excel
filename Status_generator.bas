Attribute VB_Name = "Status_generator"
Function Status()

    'Number of rows in Column A
    LastRow = Worksheets("CTC_SIL4").Range("A" & Rows.Count).End(xlUp).Row


    
    For i = 4 To LastRow

        If Worksheets("CTC_SIL4").Range("J" & i).Value = "" Then
        
        ElseIf Worksheets("CTC_SIL4").Range("K" & i).Value = "" Then
                Worksheets("CTC_SIL4").Range("L" & i).Value = "Draft"
        Else
            Worksheets("CTC_SIL4").Range("L" & i).Value = "Internally Accepted"
        End If

    Next i


End Function

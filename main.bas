Attribute VB_Name = "main"
Sub Refresh()
'''''''''''''''''''''''''''''''''''
' FILE REV,TAG,LINK AND WARNING
'

    SngStart = Timer    'Get start time.
    
    'Number of files from Column A (overall)
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    '1D Table's
    ReDim tag_cmp(LastRow)
    ReDim rev_cmp(LastRow)
    
    For i = 1 To LastRow
        rev_cmp(i) = Worksheets("CTC_SIL4").Range("J" & i).Value
        tag_cmp(i) = Worksheets("CTC_SIL4").Range("K" & i).Value
    Next i
    
    
    Trunk
    Tags
    
    
    For i = 4 To LastRow
        tag = Worksheets("CTC_SIL4").Range("K" & i).Value
        rev = Worksheets("CTC_SIL4").Range("J" & i).Value
        '14806254 - Color for OK
        '49407    - Color for Warning
        rev_bg = Worksheets("CTC_SIL4").Range("J" & i).Interior.Color


        If tag <> "" Then
            'Commit made after already tagging that file - change bg
            If rev <> rev_cmp(i) And tag = tag_cmp(i) Then
                Worksheets("CTC_SIL4").Range("J" & i).Interior.Color = 49407
            'Warning ON untill tag changes
            ElseIf rev_bg = 49407 And tag = tag_cmp(i) Then
            'Tag changes
            Else
                Worksheets("CTC_SIL4").Range("J" & i).Interior.Color = 14806254
            End If
        End If
        
        'Status
        If Worksheets("CTC_SIL4").Range("J" & i).Value = "" Then
        
        ElseIf Worksheets("CTC_SIL4").Range("K" & i).Value = "" Then
                Worksheets("CTC_SIL4").Range("L" & i).Value = "Draft"
        Else
            Worksheets("CTC_SIL4").Range("L" & i).Value = "Internally Accepted"
        End If

    Next i
    
    'Status
    
    sngEnd = Timer
    sngElapsed = sngEnd - SngStart
    Debug.Print sngElapsed

End Sub

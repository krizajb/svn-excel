Attribute VB_Name = "Tag_generator"
Function Tags()
    'Getting start time
    trunkStart = Timer
    
    Dim tagsSplit() As String
    
    'Generated file names
    Dim FileNameSheet() As String
    FileNameSheet = FileNameGenerator()
    
    Dim FileNameSVN As String
    Dim Found As Boolean
    
    
    'Number of rows in Column A
    LastRow = Worksheets("CTC_SIL4").Range("A" & Rows.Count).End(xlUp).Row
  
    'SVN address
    localhost = "10.12.7.224"
    
    tagsSheet = GetCommandOutput("cmd.exe /c svn list --depth infinity http://" & localhost & "/Project_Documentation/tags --username excel --password test")
    
    tagsSplit = Split(tagsSheet, vbCrLf)
    
    '2D Table
    'ReDim tags2D(UBound(tagsSplit), 1)
    
    'If there is something in tags
    If UBound(tagsSplit) > 1 Then
        'Delete whole tmp Worksheet
        Worksheets("tmp").Cells.Clear
        For i = 0 To UBound(tagsSplit)
            'Debug.Print tagsSplit(I)
            tmp = Split(tagsSplit(i), "/")
            'tmp(0) = tag
            'tmp(2) = file
            If UBound(tmp) = 2 Then
                If tmp(2) <> "" Then
                    For j = 0 To 1 '(UBound(tags2D, 2))
                        If j = 0 Then
                            'tags2D(i, j) = tmp(0)
                            Worksheets("tmp").Range("A" & i).Value = tmp(0)
                        Else
                            'tags2D(i, j) = tmp(2)
                            Worksheets("tmp").Range("B" & i).Value = tmp(2)
                        End If
                    Next j
                End If
            End If
        Next i
    

        'Worksheets("tmp").Select
        Worksheets("tmp").sort.SortFields.Clear
        Worksheets("tmp").sort.SortFields.Add Key:=Range("A1:A" & UBound(tagsSplit)) _
            , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With Worksheets("tmp").sort
            .SetRange Range("A1:B" & UBound(tagsSplit))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
'        Worksheets("CTC_SIL4").Range("K" & i).Select
'        selection.NumberFormat = "0.0"
    
    
   'Debug.Print UBound(FileNameSheet)

        'All files on Sheet
        For i = 1 To UBound(FileNameSheet)
            'All files on SVN server
            j = 1
            Do
                Found = False
    
                FileNameSVN = Worksheets("tmp").Range("B" & j).Value
                'Debug.Print FileNameSVN
                If FileNameSVN <> "" Then
                    If FileNameSheet(i) = FileNameSVN Then
    
                        tag = Worksheets("tmp").Range("A" & j).Value
                        If Worksheets("CTC_SIL4").Range("K" & i).Value <> tag Then
                            
       
                            Worksheets("CTC_SIL4").Range("K" & i).Value = tag
                        End If
                        Found = True
                    End If
                End If
                j = j + 1
    
            Loop Until Found Or j = UBound(tagsSplit)
    
        Next i
    
    End If
    
    'Getting end time
    trunkEnd = Timer
    'Elapsed time
    trunkElapsed = trunkEnd - trunkStart
    'Debug.Print trunkElapsed
    
End Function

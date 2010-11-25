Attribute VB_Name = "Trunk_generator"
Function Trunk()
    'Getting start time
    trunkStart = Timer

    Dim avarTmp As Variant
    Dim Found As Boolean
    
    'Generated file names
    Dim FileNames() As String
    FileNameSheet = FileNameGenerator()
    
    'Number of rows in Column A
    LastRow = Worksheets("CTC_SIL4").Range("A" & Rows.Count).End(xlUp).Row
    
    'SVN address
    localhost = "10.12.7.224"
    
    Dim trunkSplit As Variant
        
    Trunk = GetCommandOutput("cmd.exe /c svn list -v -R http://" & localhost & "/Project_Documentation/trunk/ --username excel --password test")

    'Splitting trunk, getting array of trunk log
    trunkSplit = Split(Trunk, vbCrLf)
    
    'If there is something in trunk
    If UBound(trunkSplit) > 0 Then
        'Setting our own delimiter - ";"
        For i = 0 To UBound(trunkSplit)
            Output = ""
            For j = 1 To Len(trunkSplit(i))
                ch = Mid(trunkSplit(i), j, 1)
                If ch <> " " Then
                    Output = Output & ch
                    If Mid(trunkSplit(i), j + 1, 1) = " " Then
                        Output = Output & ";"
                    End If
                End If
            Next j
            trunkSplit(i) = Output
        Next i
    
        'All files on Sheet
        For i = 4 To LastRow
            'All files on SVN server
            j = 0
            Do
                Found = False
                avarTmp = Split(trunkSplit(j), ";")
                '6 = Files
                '5 = Folders
                If UBound(avarTmp) > 5 Then
                    'avarTmp(0) = Revision
                    'avarTmp(1) = User
                    'avarTmp(3) = Month
                    'avarTmp(4) = Day
                    'avarTmp(5) = Time
                    'avarTmp(6) = Path/File
                Else: GoTo Forward
                End If
                
                tmp = Split(avarTmp(6), "/")

                If UBound(tmp) > 0 Then
                    folder = tmp(0)
                    Name = tmp(1)
                Else: GoTo Forward
                End If
                    
                'File found
                If FileNameSheet(i) = Name Then
                    Found = True
                    'Debug.Print "Found"
                    'Enter file name and revision of found file (if not there yet)
                    If Worksheets("CTC_SIL4").Range("M" & i).Value <> Name Then
                        Worksheets("CTC_SIL4").Range("M" & i) = Name
                    End If
                    
                    Worksheets("CTC_SIL4").Range("J" & i) = avarTmp(0)
                    'Enter found file hyperlink
                    path = "http://" & localhost & "/Project_Documentation/trunk/" & avarTmp(6)
                    
                    If Worksheets("CTC_SIL4").Range("N" & i).Formula <> "=HYPERLINK(""" & path & """,""" & Name & """)" Then
                        Worksheets("CTC_SIL4").Range("N" & i).Formula = "=HYPERLINK(""" & path & """,""" & Name & """)"
                        
'                        Worksheets("CTC_SIL4").Range("N" & i).Select
'                        Worksheets("CTC_SIL4").Hyperlinks.Add Anchor:=selection, Address:= _
'                        "http://" & localhost & "/Project_Documentation/trunk/" & avarTmp(6), TextToDisplay:=Name
                    End If
                'Not found
                Else
                    del = "J" & i & ":N" & i
                    'Debug.Print "Not Found"
'                    Range(del).ClearContents
                    'selection.ClearContents
                End If

Forward:
                j = j + 1
            Loop Until Found Or j = UBound(trunkSplit)
        Next i
    End If

    'Getting end time
    trunkEnd = Timer
    'Elapsed time
    trunkElapsed = trunkEnd - trunkStart
    'Debug.Print trunkElapsed
    
End Function

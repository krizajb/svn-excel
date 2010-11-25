Attribute VB_Name = "main_old"
Function AllInOneOld()
'''''''''''''''''''''''''''''''''''
' FILE REV AND LINK  + REFRESH #2
'

    SngStart = Timer    'Get start time.

    'Generated file names
    Dim FileNames() As String
    FileNameSheet = FileNameGenerator()

    'Number of rows in Column A
    LastRow = Worksheets("CTC_SIL4").Range("A" & Rows.Count).End(xlUp).Row
    
    localhost = "10.12.7.224"
    
    
    '1D Table's
    ReDim tag_cmp(LastRow)
    ReDim rev_cmp(LastRow)
    
    For i = 4 To LastRow
        rev_cmp(i) = Worksheets("CTC_SIL4").Range("J" & i).Value
        tag_cmp(i) = Worksheets("CTC_SIL4").Range("K" & i).Value
    Next i
    
    
    '2D Table
'    ReDim cmp(LastRow, 1)
'
'    'Save data before changes
'    For j = 4 To LastRow
'        For i = 0 To (UBound(cmp, 2))
'            If i = 0 Then
'                cmp(j, i) = Worksheets("CTC_SIL4").Range("J" & j).Value
'            Else
'                cmp(j, i) = Worksheets("CTC_SIL4").Range("K" & j).Value
'            End If
'        Next i
'    Next j
    
    
          
    Dim folderSplit, fileSplit, revSplit As Variant
    Dim folders, files, Output, FileNameSVN
    
    Cmd = "cmd.exe /c svn list --verbose http://" & localhost & "/documents/trunk/"
    'Folders in trunk
    folders = GetCommandOutput("cmd.exe /c svn list http://" & localhost & "/documents/trunk --username test --password test")
    'Debug.Print folders


    'Split vbCrLf - new line
    folderSplit = Split(folders, vbCrLf)
    
    Counter = 0
    'Check files in folders
    For i = 0 To (UBound(folderSplit) - 1)
        files = GetCommandOutput("cmd.exe /c svn list --verbose http://" & localhost & "/documents/trunk/" & folderSplit(i) & " --username test --password test")

        'Files in folder with details
        fileSplit = Split(files, vbCrLf)
        For j = 0 To (UBound(fileSplit))
            If fileSplit(j) <> "" Then
                Output = ""
                'We need only revision and filename
                'Seting our own delimiter - " "
                For l = 1 To Len(fileSplit(j))
                    ch = Mid(fileSplit(j), l, 1)
                    If ch <> " " Then
                        Output = Output & ch
                        If Mid(fileSplit(j), l + 1, 1) = " " Then
                            Output = Output & " "
                        End If
                    End If
                Next l

                revSplit = Split(Output, " ")
                'Avoiding index out of bound
                If UBound(revSplit) > 5 Then
                    '0 = revision
                    '6 = filename
                    
                    For k = 4 To LastRow
                        If FileNameSheet(k) <> "" Then
                            If revSplit(6) = FileNameSheet(k) Then
                                'Debug.Print revSplit(6)
                                'Counter = Counter + 1
                                
                                'Enter file name and revision of found file
                                Worksheets("CTC_SIL4").Range("J" & k) = revSplit(0)
                                Worksheets("CTC_SIL4").Range("M" & k) = revSplit(6)
                                
                                'Enter found file hyperlink
                                Worksheets("CTC_SIL4").Range("N" & k).Select
                                ActiveSheet.Hyperlinks.Add Anchor:=selection, Address:= _
                                "http://" & localhost & "/documents/trunk/" & folderSplit(i) & revSplit(6), TextToDisplay:=revSplit(6)
                            End If
                        End If
                    Next k
                End If
            End If
        Next j
    Next i
    
    'Debug.Print Counter
    
    TagDetect

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

    Next i

    sngEnd = Timer                                  ' Get end time.
    sngElapsed = format(sngEnd - SngStart, "Fixed") ' Elapsed time.
    
 
    Debug.Print sngElapsed

End Function

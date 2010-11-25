Attribute VB_Name = "Tag_generator_old"
Function TagDetect()
''''''''''''''''
' TAG DETECTOR '
''''''''''''''''
    'Get start time.
    SngStart = Timer
    
    
    'Generated file names
    Dim FileNames() As String
    FileNames = FileNameGenerator()
    
    localhost = "10.12.7.224"

    Dim tag As String
    Dim Tags
    Dim tagSplit As Variant
    
    'Number of files from Column A (overall)
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Command that will give number of tags
    Tags = GetCommandOutput("cmd.exe /c svn list http://" & localhost & "/documents/tags/")

    tagSplit = Split(Tags, vbCrLf)
    index = (UBound(tagSplit))

    ReDim Tags(UBound(tagSplit))

    For t = 0 To (UBound(tagSplit) - 1)
        tmp = Replace(tagSplit(t), "/", "")
        index = index - 1
        Tags(index) = tmp
    Next t
    

    Dim i, j
    Dim list As String
    ReDim tag_list(UBound(tagSplit))


    l = 0
    'Getting files that have tag into tag_list
    Do
        list = GetCommandOutput("cmd.exe /c svn list http://" & localhost & "/documents/tags/" & Tags(l))
        'Debug.Print list
        tag_list(l) = list
        l = l + 1
    Loop Until l = UBound(tagSplit)
    'tag_list example index=0, files with latest tag .. index=1, files with latest tag - 1

    Dim splitter As Variant
    Dim Found As Boolean
    'Checking if file has a tag
    'All files
    For j = 4 To LastRow

        If FileNames(j) <> "" Then

            i = 0 'counter for tmp - file names
            k = 0 'counter for tag_list - files with tags (svn)

            'All tags
            Do
                Found = False

                'Debug.Print i
                i = i + 1
                
                If tag_list(k) <> "" Then
                    splitter = Split(tag_list(k), vbCrLf)
                    Length = UBound(splitter)

                    l = 0
                    Do
                    'When found first tag (latest) go out of loop
                        If FileNames(j) = splitter(l) Then
                            'Debug.Print FileNames(j) & " = " & splitter(l)
                            Found = True
                            'Format of Tag Cell set to 0.0
                            Worksheets("CTC_SIL4").Range("K" & j).Select
                            selection.NumberFormat = "0.0"
                            ActiveCell.FormulaR1C1 = Tags(k)
                        End If
                        l = l + 1
                    Loop Until l = Length
                End If

                k = k + 1
            Loop Until i = UBound(tagSplit) Or Found
        End If
    Next j

    sngEnd = Timer                                  ' Get end time.
    sngElapsed = format(sngEnd - SngStart, "Fixed") ' Elapsed time.
    'Debug.Print sngElapsed

End Function

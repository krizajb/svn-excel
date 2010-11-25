Attribute VB_Name = "File_Generator"
Function FileGenerator()
'''''''''''''''
' FILE CREATE '
'''''''''''''''

    'Generated file names
    Dim FileNames() As String
    FileNames = FileNameGenerator()
    

    'Number of rows in Column A
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    'Debug.Print LastRow

    Dim avarSplit As Variant
    Dim merge
    Dim FileName As String
    
    Dim fso
    Dim word 'the Word application
    Dim doc 'the Word document
    'Dim select_ 'text selection
 
    
    Dim Test As String
    
    Set word = CreateObject("Word.application")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    
    'Till row #4 there are no files named in Sheet
    For i = 4 To LastRow
    
        'Folder name generator
        folder = Worksheets("CTC_SIL4").Range("B" & i).Value
        If folder <> "" Then
        'Debug.Print folder
            folder = StrConv(folder, vbProperCase)
            folder = Replace(folder, "Phase", "")
            folder = Replace(folder, " ", "_")
            folder = Replace(folder, "/", "_or_")
            folder = Left(folder, Len(folder) - 1)
            
        End If
        
        
        
        If FileNames(i) <> "" Then
            'Path where file is going to be saved
            file_path = "C:\Project_Documentation\" & folder & "/" & FileNames(i)

            If fso.FolderExists("C:\Project_Documentation\" & folder) = False Then
                fso.CreateFolder "C:\Project_Documentation\" & folder
            End If

            If Not fso.FileExists(file_path) Then
                Counter = Counter + 1
'                Set select_ = word.selection
'                select_.typetext "TEXT INSRETD INTO .DOCX"
                Set doc = word.documents.Add
                doc.SaveAs (file_path)

            End If
        End If
    Next i
    
    Debug.Print Counter

    word.Quit
    Set fs = Nothing
    Set word = Nothing
    
End Function

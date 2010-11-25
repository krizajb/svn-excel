Attribute VB_Name = "filename_generator"
Public Function FileNameGenerator() As String()
    start = Timer()
    
    '''''''''''''''''''''
    ' FILE NAME GENERATOR
    '
    'Generated file names from Worksheet columns
    
    Dim systemUnit As String
    Dim scope As String
    Dim FileName As String
    Dim sil As String
    
    Dim FileNames() As String
     
    
    'Number of files from Column A (overall)
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    ReDim FileNames(LastRow)
    
    Dim i As Integer
    
    For i = 4 To LastRow
        Select Case Worksheets("CTC_SIL4").Range("C" & i).Value
            Case "System"
                scope = "GEN"
                systemUnit = "SYS"

                FileName = StringReplace(scope, systemUnit, i)
                FileNames(i) = FileName
                
                FileNameGenerator = FileNames
                
                'Debug.Print FileName

            Case "Server Station"
                scope = "GEN"
                systemUnit = "SRV"

                FileName = StringReplace(scope, systemUnit, i)
                FileNames(i) = FileName
                
                FileNameGenerator = FileNames
                
                'Debug.Print FileName

            Case "Work Post Station (CCD)"
               scope = "GEN"
               systemUnit = "CCD"

              FileName = StringReplace(scope, systemUnit, i)
                FileNames(i) = FileName
                
                FileNameGenerator = FileNames
                
                'Debug.Print FileName

            Case "Remote Terminal Unit"
                scope = "GEN"
                systemUnit = "RTU"

                FileName = StringReplace(scope, systemUnit, i)
                FileNames(i) = FileName
                
                FileNameGenerator = FileNames
                
                'Debug.Print FileName

            Case "Kamnik Station Application"
                scope = "KAM"

                FileName = StringReplace(scope, systemUnit, i)
                FileNames(i) = FileName
                
                FileNameGenerator = FileNames
                
                'Debug.Print FileName

            Case Else
                FileName = ""
                FileNames(i) = FileName

                'Debug.Print FileNames(i)
        End Select
    Next i
    
    finish = Timer()
    elapsed = finish - start
    
    

    'Debug.Print elapsed

End Function

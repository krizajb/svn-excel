Attribute VB_Name = "String_replace"
Function StringReplace(scope As String, systemUnit As String, i As Integer)
    'Replaces " " ,":", "/" with "_"
    'return: File name as String
    
    Dim FileName As String
    
    FileName = Worksheets("CTC_SIL4").Range("D" & i).Value
    'Every first latter of word put into uppercase
    FileName = StrConv(FileName, vbProperCase)
    tmp = Replace(FileName, " ", "_")
    tmp2 = Replace(tmp, ":", "")
    FileName = Replace(tmp2, "/", "_or_")
    FileName = scope & "_" & FileName & "_" & systemUnit & ".docx"
       
    StringReplace = FileName
     
End Function

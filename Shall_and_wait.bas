Attribute VB_Name = "shall_and_wait"
Declare Function OpenProcess Lib "kernel32" _
                             (ByVal dwDesiredAccess As Long, _
                              ByVal bInheritHandle As Long, _
                              ByVal dwProcessId As Long) As Long

Declare Function GetExitCodeProcess Lib "kernel32" _
                                    (ByVal hProcess As Long, _
                                     lpExitCode As Long) As Long

Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103


Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub



Attribute VB_Name = "Module9"
'Module for OpenProcess, GetExitCodeProcess, CloseHandle
Private Declare Function OpenProcess Lib "kernel32" _
  (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
  (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STATUS_PENDING = &H103&

' This Sub Routine is to run any file using Shell Function and Wait until it terminates

Public Sub RunShell(cmdLine As String, mode As Byte)

    Dim hProcess As Long
    Dim ProcessId As Long
    Dim exitCode As Long

    'Execute the File
    ProcessId = Shell(cmdLine, mode)
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId)
    
    'Loop to Wait for the Execution to Terminate
    Do

        Call GetExitCodeProcess(hProcess, exitCode)
        DoEvents
   
    Loop While exitCode = STATUS_PENDING

    Call CloseHandle(hProcess)
End Sub



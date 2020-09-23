Attribute VB_Name = "Module5"
'Function to get the Drive of a give FilePath
Public Function getDrive(FilePath As String) As String
    Dim lpos As Integer
    lpos = InStr(FilePath, "\")
    getDrive = Left$(FilePath, lpos)
End Function
'Function to get the File from the FilePath
Public Function getFile(FilePath As String) As String
    Dim lpos As Integer
    lpos = InStrRev(FilePath, "\")
    lpos = Len(FilePath) - lpos
    getFile = Right$(FilePath, lpos)
End Function
'Function to get the path from the FilePath
Public Function getPath(FilePath As String) As String
    Dim lpos As Integer
    lpos = InStrRev(FilePath, "\") - 1
    getPath = Left$(FilePath, lpos)
End Function
'Procedure to change the Directory
Public Sub changeDirectory(ByVal Drive As String, ByVal path As String)
    ChDrive Drive
    ChDir path
End Sub
'Procedure to the application path
Public Sub changeapppath()
    ChDrive App.path   ' Set the drive.
    ChDir App.path      ' Set the directory.
End Sub
'Function to remove the Extension from the file.
Public Function removeExt(FileName As String) As String
    Dim lpos As Integer
    lpos = InStr(FileName, ".") - 1
    removeExt = Left$(FileName, lpos)
End Function

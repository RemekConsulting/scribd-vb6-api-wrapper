VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- This object provides basic File Management functions that are required by applications.
'---
'--- AUTHOR:    Greg Bridle
'--- DATE:      2001.10.04.
'---
'--- PATCH HISTORY
'---
'--- DATE       BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

Event LogError(ByVal Procedure As String, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nstrBuffererLength As Long, ByVal lpstrBufferer As String) As Long
   
'-----------------------------------------------------------
'--- P U B L I C   R O U T I N E S   D E F I N E D   H E R E
'-----------------------------------------------------------
Public Function GetFileName(CheckFile) As String

    On Error GoTo GetFileName_Error
    
    Dim strFileName             As String
    
    GetFileName = CheckFile
    
    If Not InStrRev(CheckFile, "\") = 0 Then GetFileName = Mid$(CheckFile, InStrRev(CheckFile, "\") + 1)
    
GetFileName_Error:

    Select Case Err.Number
    Case 0
    Case Else
        RaiseEvent LogError("GetFileName", Err.Number, Err.Description)
    End Select
    
End Function

Public Function GetFileType(CheckFile) As String

    On Error GoTo GetFileName_Error
    
    Dim strFileName             As String
    
    GetFileType = CheckFile
    
    If Not InStrRev(CheckFile, ".") = 0 Then GetFileType = Mid$(CheckFile, InStrRev(CheckFile, ".") + 1)
    
GetFileName_Error:

    Select Case Err.Number
    Case 0
    Case Else
        RaiseEvent LogError("GetFileType", Err.Number, Err.Description)
    End Select
    
End Function

Public Function GetDirectory(FileName As String) As String

    On Error GoTo ErrorHandler

    '--- Extract the directory name from the file name.
    If InStrRev(FileName, "\") Then
        GetDirectory = Left$(FileName, InStrRev(FileName, "\") - 1)
    Else
        GetDirectory = FileName
    End If

ErrorHandler:
    
    Select Case Err.Number
    Case 0
    Case Else
        RaiseEvent LogError("GetDirectory", Err.Number, Err.Description)
    End Select
    
    Exit Function
    Resume
    
End Function

Public Function GetParentNode(URL As String) As String

    On Error GoTo ErrorHandler

    '--- Extract the directory name from the file name.
    If InStrRev(URL, "/") Then
        GetParentNode = Left$(URL, InStrRev(URL, "/") - 1)
    Else
        GetParentNode = URL
    End If

ErrorHandler:
    
    Select Case Err.Number
    Case 0
    Case Else
        RaiseEvent LogError("GetParentNode", Err.Number, Err.Description)
    End Select
    
    Exit Function
    Resume
    
End Function

Public Function GetParentDomain(URL As String) As String

    On Error GoTo ErrorHandler

    Dim strWorker                       As String
    
    '--- Extract the main domain from a URL
    If InStrRev(URL, "http://") Then
        strWorker = Mid$(URL, 8)
        If InStr(strWorker, "/") Then
            GetParentDomain = "http://" & Left$(strWorker, InStr(strWorker, "/") - 1)
        Else
            GetParentDomain = strWorker
        End If
    Else
        GetParentDomain = Left$(URL, InStr(URL, "/") - 1)
    End If

ErrorHandler:
    
    Select Case Err.Number
    Case 0
    Case Else
        RaiseEvent LogError("GetParentDomain", Err.Number, Err.Description)
    End Select
    
    Exit Function
    Resume
    
End Function

Public Function DirectoryCreate(DirectoryName As String) As Boolean

    Dim strWorkDirectory                As String
    Dim strNewDirectory                 As String

    DirectoryCreate = False
    
    On Error GoTo ErrorHandler

    '--- First we have to check if the drive supplied is actually valid. And if not then we exit the whole process
    If Not IsValidDrive(Left$(DirectoryName, 1)) Then
        Exit Function
    End If
    
    '--- Don't try and create a directory if none is supplied
    If Not Len(DirectoryName) > 0 Then
        DirectoryCreate = True
        Exit Function
    End If
    
    '--- Attempt to create the whole directory first off
    strWorkDirectory = DirectoryName
    strNewDirectory = strWorkDirectory
    MkDir strWorkDirectory
    
    GoTo ErrorHandler

DirectoryCreate_Retry:
    
    '--- Strip back the directory name to find the last valid level of directory
    Do
        Do
            strNewDirectory = GetDirectory(strNewDirectory)
            If Not Len(strNewDirectory) < 4 Then
                MkDir strNewDirectory
            End If
            Exit Do
        Loop
        strNewDirectory = strWorkDirectory
        MkDir strWorkDirectory
        Exit Do
    Loop

ErrorHandler:
    
    Select Case Err.Number
    Case 0
        DirectoryCreate = True
    '--- Directory already exists
    Case 75
        DirectoryCreate = True
    Case 76
        Resume DirectoryCreate_Retry
    Case Else
        RaiseEvent LogError("DirectoryCreate", Err.Number, Err.Description)
    End Select
    
    Exit Function

End Function

Public Function RemoveLastDirNode(DirectoryPath As String) As String

    If Right$(DirectoryPath, 1) = "\" Then
        DirectoryPath = Left$(DirectoryPath, Len(DirectoryPath) - 1)
    End If
    
    If Not InStrRev(DirectoryPath, "\") = 0 Then
        RemoveLastDirNode = Left$(DirectoryPath, InStrRev(DirectoryPath, "\"))
    Else
        RemoveLastDirNode = DirectoryPath
    End If

End Function

Public Function GetLastDirNode(ByVal DirectoryPath) As String

    If Right$(DirectoryPath, 1) = "\" Then
        DirectoryPath = Left$(DirectoryPath, Len(DirectoryPath) - 1)
    End If
    
    If Not InStrRev(DirectoryPath, "\") = 0 Then
        GetLastDirNode = Mid$(DirectoryPath, InStrRev(DirectoryPath, "\") + 1)
    Else
        GetLastDirNode = DirectoryPath
    End If

End Function

Public Function IsValidDrive(DriveLetter As String) As Boolean

    Dim strBuffer                       As String
    Dim lngBufferSize                   As Long
   
    '--- Call the API with a strBufferer size of 0. The call fails, and the required size
    '--- is returned as the result.
    lngBufferSize = GetLogicalDriveStrings(0&, strBuffer)
    
    '--- Pad a strBufferer to hold the results
    strBuffer = Space$(lngBufferSize)
    lngBufferSize = Len(strBuffer)
    
    '--- And call again
    If GetLogicalDriveStrings(lngBufferSize, strBuffer) Then
    
        '--- If the drive letter passed is in the returned logical drive string,
        '--- return True.
        IsValidDrive = InStr(1, strBuffer, DriveLetter, vbTextCompare) > 0
    
    End If
    
End Function


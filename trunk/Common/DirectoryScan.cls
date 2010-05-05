VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- This object is used to scan through a directory. You can specify a file type and/or attribute to
'--- search for and only those files will be processed.
'---
'--- AUTHOR: Greg Bridle
'--- DATE:   2001.10.04.
'---
'--- PATCH HISTORY
'---
'--- DATE       BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'
'--- 2008.11.24 Greg        Added inclusion and exclusion search patterns

Option Explicit

'--- Property variables
Private propWalkDirectories()           As String
Private propWalkFileDir()               As String

Private Type WalkFile
    Name                                As String
    FileName                            As String
    Directory                           As String
    DateTime                            As Double
    Attributes                          As String
    Permissions                         As String
    FileSize                            As Long
End Type
Private propWalkFile()                  As WalkFile

Private propWalkDirectoriesCount        As Long
Private propWalkFilesCount              As Long
Private propExtendedAttributes          As Boolean
Private propCancel                      As Boolean
Private propPauseState                  As Boolean
Private propExcludeFilesContaining      As String
Private propIncludeFilesContaining      As String
Private propDirectoriesOnly             As Boolean

Event WalkingDirectory(DirectoryName)
Event WalkingFile(DirectoryName As String, FileName As String)

'--- Working variables
Private arrayExcludeFilesContaining()   As String
Private arrayIncludeFilesContaining()   As String
Private intExcludeFilesContaining       As Integer
Private intIncludeFilesContaining       As Integer

Private Const STATUS_TIMEOUT = &H102&
Private Const INFINITE = -1& ' Infinite interval
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'-----------------------------------------------------------
'--- P U B L I C   R O U T I N E S   D E F I N E D   H E R E
'-----------------------------------------------------------
Public Sub WalkDirectory(ByVal DirectoryName As String, ByVal ScanSubDirs As Boolean, ByVal FileType As String, ByVal Attributes As Integer, Optional KeepPrevious As Boolean = False, Optional ExcludeTypes As String = "")

    Dim strDirectoryName                As String
    
    propCancel = False
    
    If Not KeepPrevious Then
        
        Erase propWalkDirectories
        Erase propWalkFile
        
        propWalkDirectoriesCount = 0
        propWalkFilesCount = 0
    
    Else
        
        If Not propWalkFilesCount > 0 Then
            propWalkFilesCount = 0
        End If
    
    End If
    
    If Not Right$(DirectoryName, 1) = "\" Then
        strDirectoryName = DirectoryName & "\"
    Else
        strDirectoryName = DirectoryName
    End If
        
    WalkDirectoryTree strDirectoryName, ScanSubDirs, FileType, Attributes, ExcludeTypes
    
    propWalkFilesCount = propWalkFilesCount - 1

End Sub

'--- This routine is used when the software is shutting down but the directory walk is still in
'--- progress.
Public Sub Cancel()
    propCancel = True
End Sub

'--- These routines are used when the application needs to pause for any reason. It stops the directory
'--- scan from continuing until the application is unpaused.
Public Sub Pause()
    propPauseState = True
End Sub

Public Sub Unpause()
    propPauseState = False
End Sub

'-------------------------------------------------------------
'--- P R I V A T E   R O U T I N E S   D E F I N E D   H E R E
'-------------------------------------------------------------
Private Sub WalkDirectoryTree(ByVal DirName, ByVal ScanSubDirs As Boolean, Optional FileType As String, Optional Attributes As Integer, Optional ExcludeTypes As String = "")

    Dim strFileName                     As String
    Dim strFileType                     As String
    Dim strDriveName                    As String
    Dim strDirs()                       As String
    Dim intDirCount                     As Integer
    Dim intPointer                      As Integer
    
    Dim intLoopCount                    As Integer
    
    On Error GoTo ErrorHandler
    
    '--- If the user has cancelled the directory walk then we just exit straight away
    If propCancel Then Exit Sub
    
    RaiseEvent WalkingDirectory(DirName)
    
    DoEvents
    
    strFileName = Dir(DirName, vbDirectory Or vbReadOnly Or vbSystem Or vbHidden)
        
    Do While strFileName <> ""
        
        If strFileName <> "." And strFileName <> ".." Then
            
            If (GetAttr(DirName & strFileName) And vbDirectory) = vbDirectory Then
                
                intDirCount = intDirCount + 1
                ReDim Preserve strDirs(intDirCount) As String
                strDirs(intDirCount) = DirName & strFileName & "\"

                ReDim Preserve propWalkDirectories(propWalkDirectoriesCount) As String
                propWalkDirectories(propWalkDirectoriesCount) = DirName & strFileName & "\"
                propWalkDirectoriesCount = propWalkDirectoriesCount + 1
            
            ElseIf Not propDirectoriesOnly Then
            
                '--- If included types are selected then we only include those files
                '--- with an extension which is in our list.
                If Not Len(FileType) = 0 Then
                    
                    If Not InStrRev(strFileName, ".") = 0 Then
                        strFileType = LCase$(Mid$(strFileName, InStrRev(strFileName, ".") + 1))
                        If Not InStr(1, FileType, strFileType, vbTextCompare) > 0 Then
                            GoTo NextFile
                        End If
                    Else
                        GoTo NextFile
                    End If
                
                End If
                
                '--- If necessary exlude the selected file types from the transfer
                If Not Len(ExcludeTypes) = 0 Then
                    
                    If Not InStrRev(strFileName, ".") = 0 Then
                        strFileType = LCase$(Mid$(strFileName, InStrRev(strFileName, ".") + 1))
                        If InStr(1, ExcludeTypes, strFileType, vbTextCompare) > 0 Then
                            GoTo NextFile
                        End If
                    End If
                
                End If
                
                If Not Len(propExcludeFilesContaining) = 0 Then
                    For intExcludeFilesContaining = 0 To UBound(arrayExcludeFilesContaining)
                        If InStr(1, DirName & strFileName, arrayExcludeFilesContaining(intExcludeFilesContaining), vbTextCompare) > 0 Then
                            GoTo NextFile
                        End If
                    Next intExcludeFilesContaining
                End If
                
                If Not Len(propIncludeFilesContaining) = 0 Then
                    For intIncludeFilesContaining = 0 To UBound(arrayIncludeFilesContaining)
                        If Not InStr(1, DirName & strFileName, arrayIncludeFilesContaining(intIncludeFilesContaining), vbTextCompare) > 0 Then
                            GoTo NextFile
                        End If
                    Next intIncludeFilesContaining
                End If
                
                If Not Attributes = 0 Then
                    If Not (GetAttr(DirName & strFileName) And Attributes) = Attributes Then
                        GoTo NextFile
                    End If
                End If
                
                RaiseEvent WalkingFile(CStr(DirName), strFileName)
                
                ReDim Preserve propWalkFile(propWalkFilesCount) As WalkFile

                With propWalkFile(propWalkFilesCount)
                    
                    .Directory = DirName
                    .Name = DirName & strFileName
                    .FileName = strFileName
                    .DateTime = FileDateTime(.Name)
                    .FileSize = FileLen(.Name)
                        
                    If propExtendedAttributes Then
                        
                        If (GetAttr(DirName & strFileName) And vbReadOnly) = vbReadOnly Then
                            .Attributes = .Attributes & "R"
                        Else
                            .Attributes = .Attributes & "-"
                        End If
                        If (GetAttr(DirName & strFileName) And vbHidden) = vbHidden Then
                            .Attributes = .Attributes & "H"
                        Else
                            .Attributes = .Attributes & "-"
                        End If
                        If (GetAttr(DirName & strFileName) And vbSystem) = vbSystem Then
                            .Attributes = .Attributes & "S"
                        Else
                            .Attributes = .Attributes & "-"
                        End If
                        If (GetAttr(DirName & strFileName) And vbArchive) = vbArchive Then
                            .Attributes = .Attributes & "A"
                        Else
                            .Attributes = .Attributes & "-"
                        End If
                    End If
                
                End With
                
                propWalkFilesCount = propWalkFilesCount + 1
                
                If propWalkFilesCount Mod 20 = 0 Then
                    DoEvents
                End If
                
                '--- If the user has cancelled the directory walk then we just exit straight away
                If propCancel Then Exit Sub

NextFile:
            
            End If
            
        End If
        
    
        '--- If the application is paused then we wait until it is unpaused before we continue
        While propPauseState
            Wait 100
        Wend

        strFileName = Dir
        
        intLoopCount = intLoopCount + 1
        If Not intLoopCount < 100 Then
            intLoopCount = 0
            DoEvents
        End If
        
    Loop
    
    '--- If this directory has sub-directories then we'll walk those too
    If intDirCount > 0 And ScanSubDirs Then
        
        For intPointer = 1 To intDirCount
            
            WalkDirectoryTree strDirs(intPointer), ScanSubDirs, FileType, Attributes, ExcludeTypes
            
            '--- If the user has cancelled the directory walk then we just exit straight away
            If propCancel Then Exit Sub
    
        Next intPointer
    
    End If

ErrorHandler:

    Select Case Err.Number
    Case 0
    Case 5, 53
        Resume NextFile
    Case Else
    End Select
    
    Exit Sub
    Resume
    
End Sub

'--- The MsgWaitObj function replaces Sleep, WaitForSingleObject, WaitForMultipleObjects functions.
'--- Unlike these functions, it doesn't block thread messages processing.
'---
'--- Using instead Sleep: MsgWaitObj dwMilliseconds
'--- Using instead WaitForSingleObject: retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
'--- Using instead WaitForMultipleObjects: retval = MsgWaitObj(dwMilliseconds, hObj(0&), n)
'---                                 where n      - wait objects quantity,
'---                                       hObj() - their handles array.
Public Function Wait(Interval As Long, Optional hObj As Long = 0&, Optional nObj As Long = 0&) As Long

    Dim lngTimer1                           As Long
    Dim lngTimer2                           As Long
    
    If Not Interval = INFINITE Then
    
        lngTimer1 = GetTickCount()
        
        On Error Resume Next
        
        lngTimer1 = lngTimer1 + Interval
        '--- Overflow prevention
        If Not Err = 0& Then
            If lngTimer1 > 0& Then
                lngTimer1 = ((lngTimer1 + &H80000000) + Interval) + &H80000000
            Else
                lngTimer1 = ((lngTimer1 - &H80000000) + Interval) - &H80000000
            End If
        End If
        On Error GoTo 0
        '--- lngTimer1 contains now absolute time of the end of interval
    
    Else
        
        lngTimer2 = INFINITE
    
    End If

    Do
    
        If Interval <> INFINITE Then
        
            lngTimer2 = GetTickCount()
            
            On Error Resume Next
            
             lngTimer2 = lngTimer1 - lngTimer2
            '--- Overflow prevention
            If Err <> 0& Then
                If lngTimer1 > 0& Then
                    lngTimer2 = ((lngTimer1 + &H80000000) - (lngTimer2 - &H80000000))
                Else
                    lngTimer2 = ((lngTimer1 - &H80000000) - (lngTimer2 + &H80000000))
                End If
            End If
            On Error GoTo 0
            
            '--- lngTimer2 contains now the remaining interval part
            If IIf((lngTimer2 Xor Interval) > 0&, _
                lngTimer2 > Interval, lngTimer2 < 0&) Then
                '--- Interval expired
                '--- during DoEvents
                Wait = STATUS_TIMEOUT
                Exit Function
            End If
            
        End If
        
        '--- Wait for event, interval expiration
        '--- or message appearance in thread queue
        Wait = MsgWaitForMultipleObjects(nObj, hObj, 0&, lngTimer2, QS_ALLINPUT)
        '--- Let's message be processed
        DoEvents
        If Wait <> nObj Then
            Exit Function
        End If
        '--- It was message - continue to wait
        
    Loop

End Function

'-----------------------------------------------------------
'--- P R O P E R T Y   V A L U E S   A R E   S E T   H E R E
'-----------------------------------------------------------
Public Property Get WalkDirectories(Index As Long) As String
    WalkDirectories = propWalkDirectories(Index)
End Property
Public Property Get WalkFile(Index As Long) As String
    WalkFile = propWalkFile(Index).Name
End Property
Public Property Get WalkFileDir(Index As Long) As String
    WalkFileDir = propWalkFile(Index).Directory
End Property
Public Property Get WalkFileName(Index As Long) As String
    WalkFileName = propWalkFile(Index).FileName
End Property
Public Property Get WalkFileDateTime(Index As Long) As Double
    WalkFileDateTime = propWalkFile(Index).DateTime
End Property
Public Property Get WalkFileSize(Index As Long) As Long
    WalkFileSize = propWalkFile(Index).FileSize
End Property
Public Property Get WalkFileAttributes(Index As Long) As String
    WalkFileAttributes = propWalkFile(Index).Attributes
End Property
Public Property Get WalkDirectoriesCount() As Long
    WalkDirectoriesCount = propWalkDirectoriesCount
End Property
Public Property Get WalkFilesCount() As Long
    WalkFilesCount = propWalkFilesCount
End Property

Public Property Let ExtendedAttributes(vData As Boolean)
    propExtendedAttributes = vData
End Property
Public Property Let ExcludeFilesContaining(vData As String)
    propExcludeFilesContaining = vData
    arrayExcludeFilesContaining = Split(vData, "|")
End Property
Public Property Let IncludeFilesContaining(vData As String)
    propIncludeFilesContaining = vData
    arrayIncludeFilesContaining = Split(vData, "|")
End Property

'-------------------------------------------------------------------
'--- C L A S S   I N I T I A L I Z A T I O N   O C C U R S   H E R E
'-------------------------------------------------------------------
Private Sub Class_Initialize()
    propExtendedAttributes = False
End Sub

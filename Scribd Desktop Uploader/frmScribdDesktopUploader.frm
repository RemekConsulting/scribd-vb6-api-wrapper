VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmScribdDesktopUploader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scribd. Desktop Uploader"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   Icon            =   "frmScribdDesktopUploader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDisallowPrintingAndSelection 
      Caption         =   "Disallow P&rinting and Text Selection"
      Height          =   405
      Left            =   3300
      TabIndex        =   21
      Top             =   2460
      Width           =   3795
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   7005
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   18309
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkPrivate 
      Caption         =   "Upload documents as &Private"
      Height          =   405
      Left            =   90
      TabIndex        =   7
      Top             =   2460
      Width           =   3165
   End
   Begin VB.TextBox txtUserId 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtSessionKey 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   7680
      TabIndex        =   4
      Top             =   1110
      Width           =   2325
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   345
      Left            =   9390
      TabIndex        =   6
      Top             =   2070
      Width           =   825
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   525
      Left            =   1380
      TabIndex        =   9
      Top             =   2970
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   8
      Top             =   2970
      Width           =   1245
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   2070
      Width           =   9255
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtAPIKey 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   330
      Width           =   2325
   End
   Begin VB.TextBox txtInputBuffer 
      Height          =   2865
      Left            =   5130
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   4080
      Width           =   5205
   End
   Begin VB.TextBox txtOutputBuffer 
      Height          =   2865
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   4050
      Width           =   5025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   11000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   0
      X2              =   11000
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   5160
      TabIndex        =   19
      Top             =   870
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Session Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   7680
      TabIndex        =   18
      Top             =   870
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   11000
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   11000
      Y1              =   3630
      Y2              =   3630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Directory to upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   90
      TabIndex        =   17
      Top             =   1830
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   16
      Top             =   870
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   870
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "API Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   90
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Incoming HTTP Response"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   3840
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Outgoing HTTP Request"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   12
      Top             =   3840
      Width           =   2100
   End
End
Attribute VB_Name = "frmScribdDesktopUploader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- A basic application which allows the user to browse for a folder and then upload all the documents
'--- found in that folder.
'---
'--- AUTHOR:        Greg Bridle
'--- DATE:          2010.04.16
'---
'--- DEPENDANCIES:  Scribd_Document.cls
'---                Microsoft XML 6.0
'---
'--- PATCH HISTORY
'---
'--- DATE           BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

Private WithEvents objScribd            As Scribd_VB6
Attribute objScribd.VB_VarHelpID = -1
Private objLastPosition                 As New LastPosition
Private objLastValue                    As New LastValue
Private objTextBox                      As New TextBoxClass
Private objFileManagement               As New FileManagement
Private objDirectoryScan                As New DirectoryScan

Private lngPointer                      As Long

'--- Happens when the user clicks on the Browse button. They will be shown a dialog
'--- where they can select the folder they want to upload documents from.
Private Sub cmdBrowse_Click()

    txtFileName = BrowseForFolder(Me.hWnd, , "Folder to upload to Scribd")

End Sub

'--- When the user clicks on the Ok button then we first attempt a login to Scribd. If
'--- that is successful then we walk the directory looking for files to upload and
'--- attempt to upload them to Scribd one by one.
Private Sub cmdOk_Click()

    Dim lngFilePointer                  As Long
    
    cmdOk.Enabled = False
    
    '--- Clear any previous results
    txtOutputBuffer.Text = ""
    txtInputBuffer.Text = ""

    With objScribd
    
        .APIKey = txtAPIKey
        .Username = txtUsername
        .Password = txtPassword
        .SessionKey = txtSessionKey
        
        .Method = "user.login"
        
        '--- Only continue if the login is successful
        If .DoAction Then
        
            txtUserId = .UserId
            txtSessionKey = .SessionKey
            
            '--- Loop through the selected directory and start looking for documents to upload
            objDirectoryScan.WalkDirectory txtFileName, True, "", 0
        
            '--- Loop through all files found by the Directory Scan object
            For lngFilePointer = 0 To objDirectoryScan.WalkFilesCount
                    
                .Method = "docs.upload"
                .MimeType = "application/octet-stream"
                .CreateDocument
                .Document.FileName = objDirectoryScan.WalkFile(lngFilePointer)
                .Document.DocumentType = objFileManagement.GetFileType(objDirectoryScan.WalkFile(lngFilePointer))
                
                '--- If the document should be private then we set that flag
                If chkPrivate.Value = 1 Then
                    .Document.Access = scribdPrivateDocument
                Else
                    .Document.Access = scribdPublicDocument
                End If
                
                '--- Now upload the document to Scribd
                If .DoAction Then
                
                    '--- If the document upload was successful then we display the received buffer
                    txtInputBuffer.Text = objScribd.Result
                    
                    '--- If the user is going to disallow Printing and Text Selection then we do that here
                    If chkDisallowPrintingAndSelection.Value = 1 Then
                        ChangeSettings objScribd.Document.DocumentId
                    End If
                
                Else
                    '--- Notify the user of any issues with the upload
                    MsgBox "Method docs.upload in the Scribd API failed for the document '" & objDirectoryScan.WalkFile(lngFilePointer) & "' with the following error: " & .ErrorNumber & " - " & .ErrorDescription, vbCritical + vbOKOnly, "Scribd Desktop Uploader"
                    LogError "Method docs.upload in the Scribd API failed for the document '" & objDirectoryScan.WalkFile(lngFilePointer) & "' with the following error: " & .ErrorNumber & " - " & .ErrorDescription
                    txtInputBuffer.Text = objScribd.Result
                End If
                
            Next lngFilePointer
            
        Else
            '--- Notify the user of any issues with the login
            MsgBox "Method user.login in the Scribd API failed for the document '" & objDirectoryScan.WalkFile(lngFilePointer) & "' with the following error: " & .ErrorNumber & " - " & .ErrorDescription, vbCritical + vbOKOnly, "Scribd Desktop Uploader"
            LogError "Method user.login in the Scribd API failed for the document '" & objDirectoryScan.WalkFile(lngFilePointer) & "' with the following error: " & .ErrorNumber & " - " & .ErrorDescription
            txtInputBuffer.Text = objScribd.Result
        End If
            
    End With
    
    cmdOk.Enabled = True
    
End Sub

'--- This function is called to change document settings after the document has been uploaded
'--- to Scribd.
Private Function ChangeSettings(DocumentId As Long)

    With objScribd
        .Method = "docs.changeSettings"
        .CreateDocument
        .Document.DocumentId = DocumentId
        .Document.DisablePrint = (chkPrivate.Value = 1)
        .Document.DisableSelectText = (chkDisallowPrintingAndSelection.Value = 1)
        If Not .DoAction Then
            LogError "Method docs.changeSettings in the Scribd API failed for Document Id " & DocumentId & " with the following error: " & .ErrorNumber & " - " & .ErrorDescription
        End If
    End With
    
End Function

Private Sub cmdExit_Click()

    cmdExit.Enabled = False
    Unload Me
    
End Sub

'--- This event is raised by the Scribd Object whenever it sends data to Scribd
Private Sub objScribd_PostingData(DataToPost As Variant)
    txtOutputBuffer.Text = txtOutputBuffer.Text & DataToPost
End Sub

'--- Log any errors to a log file so that we have a track of what happened and when
Private Sub LogError(LogData As String)

    Dim intLogFile                      As Integer
    
    intLogFile = FreeFile
    Open App.Path & "\debug.log" For Append As #intLogFile
    Print #intLogFile, LogData
    Close #intLogFile

    StatusBar1.Panels(1).Text = LogData
    
End Sub

'--- When the form loads we restore the position and last set values. When it closes
'--- down then we save the position and last set values.
Private Sub Form_Load()

    Set objScribd = New Scribd_VB6
    
    With objLastValue
        .GetLastValue App.EXEName, txtAPIKey
        .GetLastValue App.EXEName, txtUsername
        .GetLastValue App.EXEName, txtPassword
        .GetLastValue App.EXEName, txtFileName
        .GetLastValue App.EXEName, chkPrivate
        .GetLastValue App.EXEName, chkDisallowPrintingAndSelection
        .GetLastValue App.EXEName, txtOutputBuffer
        .GetLastValue App.EXEName, txtInputBuffer
    End With
    
    objLastPosition.GetLastPosition App.EXEName, "Main", Me, True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objScribd = Nothing
    
    With objLastValue
        .SaveLastValue App.EXEName, txtAPIKey
        .SaveLastValue App.EXEName, txtUsername
        .SaveLastValue App.EXEName, txtPassword
        .SaveLastValue App.EXEName, txtFileName
        .SaveLastValue App.EXEName, chkPrivate
        .SaveLastValue App.EXEName, chkDisallowPrintingAndSelection
        .SaveLastValue App.EXEName, txtOutputBuffer
        .SaveLastValue App.EXEName, txtInputBuffer
    End With
    
    objLastPosition.SaveLastPosition App.EXEName, "Main", Me

End Sub

'--- These events happen when the textbox controls gain focus
Private Sub txtAPIKey_GotFocus()
    objTextBox.AutoSelect txtAPIKey
End Sub

Private Sub txtFileName_GotFocus()
    objTextBox.AutoSelect txtFileName
End Sub

Private Sub txtInputBuffer_GotFocus()
    objTextBox.AutoSelect txtInputBuffer
End Sub

Private Sub txtOutputBuffer_GotFocus()
    objTextBox.AutoSelect txtOutputBuffer
End Sub

Private Sub txtPassword_GotFocus()
    objTextBox.AutoSelect txtPassword
End Sub

Private Sub txtSessionKey_GotFocus()
    objTextBox.AutoSelect txtSessionKey
End Sub

Private Sub txtUserId_GotFocus()
    objTextBox.AutoSelect txtUserId
End Sub

Private Sub txtUsername_GotFocus()
    objTextBox.AutoSelect txtUsername
End Sub

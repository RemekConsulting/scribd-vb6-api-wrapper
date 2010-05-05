VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmScribdDocumentUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scribd. Document Update"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   Icon            =   "frmScribdDocumentUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutputBuffer 
      Height          =   2865
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   3300
      Width           =   5025
   End
   Begin VB.TextBox txtInputBuffer 
      Height          =   2865
      Left            =   5130
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   3330
      Width           =   5205
   End
   Begin VB.TextBox txtAPIKey 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   330
      Width           =   2325
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1110
      Width           =   2325
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   7
      Top             =   2220
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   525
      Left            =   1380
      TabIndex        =   8
      Top             =   2220
      Width           =   1245
   End
   Begin VB.TextBox txtSessionKey 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   7680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtUserId 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1110
      Width           =   2325
   End
   Begin VB.CheckBox chkPrivate 
      Caption         =   "Mark documents as &Private"
      Height          =   405
      Left            =   90
      TabIndex        =   5
      Top             =   1710
      Width           =   3165
   End
   Begin VB.CheckBox chkDisallowPrintingAndSelection 
      Caption         =   "Disallow P&rinting and Text Selection"
      Height          =   405
      Left            =   3300
      TabIndex        =   6
      Top             =   1710
      Width           =   3795
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   6375
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   18
      Top             =   3090
      Width           =   2100
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
      TabIndex        =   17
      Top             =   3090
      Width           =   2235
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
      TabIndex        =   16
      Top             =   90
      Width           =   690
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
      TabIndex        =   14
      Top             =   870
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   11000
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   11000
      Y1              =   2910
      Y2              =   2910
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
      TabIndex        =   13
      Top             =   870
      Width           =   1050
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
      TabIndex        =   12
      Top             =   870
      Width           =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   0
      X2              =   11000
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   11000
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmScribdDocumentUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- Implements a wrapper for the Scrbid API to be used with VB6 (The forgotten programming language).
'---
'--- AUTHOR:    Greg Bridle
'--- DATE:      2010.04.16
'---
'--- PATCH HISTORY
'---
'--- DATE       BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

Private WithEvents objScribd            As Scribd_VB6
Attribute objScribd.VB_VarHelpID = -1
Private objLastPosition                 As New LastPosition
Private objLastValue                    As New LastValue
Private objTextBox                      As New TextBoxClass

'--- When the user clicks on the Ok button then we first attempt a login to Scribd. If
'--- that is successful then we request a list of files from Scribd in blocks of 10
'--- and attempt to update the document settings on Scribd one by one.
Private Sub cmdOk_Click()

    Dim lngDocumentOffset               As Long
    Dim intPointer                      As Integer
    
    cmdOk.Enabled = False
    
    '--- Clear any previous results
    txtInputBuffer.Text = ""
    txtOutputBuffer.Text = ""

    With objScribd
    
        .APIKey = txtAPIKey
        .Username = txtUsername
        .Password = txtPassword
        .SessionKey = txtSessionKey
        
        .Method = "user.login"
        
        '--- Only continue if the login is successful
        If .DoAction Then
        
            Do
            
                .Method = "docs.getList"
                .DocumentLimit = 10
                .DocumentOffset = lngDocumentOffset
            
Retry_GetDocuments:
            
                If .DoAction Then
                
                    If Not .DocumentCount = 0 Then
                        
                        For intPointer = 0 To .DocumentCount - 1
                        
                            .Method = "docs.changeSettings"
                            .DocumentNumber = intPointer
                            
                            '--- Set the new document attributes
                            .Document.Author = "My Document Author"
                            .Document.Publisher = "My Document Publisher"
                            .Document.License = scribdLicense_Copyright

                            '--- If requested make the document private
                            If (chkPrivate.Value = 1) Then
                                .Document.Access = scribdPrivateDocument
                            Else
                                .Document.Access = scribdPublicDocument
                            End If
                                
                            '--- Set the download formats to None and if requested disable
                            '--- printing and text selection
                            .Document.DownloadFormats = scribdDownload_None
                            .Document.DisablePrint = (chkDisallowPrintingAndSelection.Value = 1)
                            .Document.DisableSelectText = (chkDisallowPrintingAndSelection.Value = 1)
                            
                            ShowDebug "Processing document number " & CStr(lngDocumentOffset + intPointer) & " '" & .Document.Title & "'."
Retry_Processing:
                            
                            If Not .DoAction Then
                                If MsgBox("Method docs.changeSetting on Scribd API failed with the following error: " & .ErrorNumber & " - " & .ErrorDescription, vbCritical + vbRetryCancel, "Scribd Document Update") = vbRetry Then
                                    GoTo Retry_Processing
                                End If
                            End If
                            
                            txtInputBuffer.Text = .Result
                            DoEvents
            
                        Next intPointer
                        
                    Else
                        Exit Do
                    End If
                
                Else
                    If MsgBox("Method docs.getList on Scribd API failed with the following error: " & .ErrorNumber & " - " & .ErrorDescription, , vbCritical + vbRetryCancel, "Scribd Document Update") = vbRetry Then
                        GoTo Retry_GetDocuments
                    End If
                End If
                
                lngDocumentOffset = lngDocumentOffset + .DocumentCount
                
            Loop
            
        Else
            '--- Notify the user of any issues with the login
            MsgBox "Method user.login in the Scribd API failed with the following error: " & .ErrorNumber & " - " & .ErrorDescription, vbCritical + vbOKOnly, "Scribd Document Update"
            LogError "Method user.login in the Scribd API failed with the following error: " & .ErrorNumber & " - " & .ErrorDescription
            txtInputBuffer.Text = objScribd.Result
        End If
    
    End With
    
    cmdOk.Enabled = True
    cmdExit.Enabled = True

End Sub

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

Private Sub ShowDebug(DebugMessage As String)
    StatusBar1.Panels(1).Text = DebugMessage
End Sub

'--- When the form loads we restore the position and last set values. When it closes
'--- down then we save the position and last set values.
Private Sub Form_Load()

    Set objScribd = New Scribd_VB6
    
    With objLastValue
        .GetLastValue App.EXEName, txtAPIKey
        .GetLastValue App.EXEName, txtUsername
        .GetLastValue App.EXEName, txtPassword
        .GetLastValue App.EXEName, chkPrivate
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
        .SaveLastValue App.EXEName, chkPrivate
        .SaveLastValue App.EXEName, txtOutputBuffer
        .SaveLastValue App.EXEName, txtInputBuffer
    End With
    
    objLastPosition.SaveLastPosition App.EXEName, "Main", Me

End Sub

'--- These events happen when the textbox controls gain focus
Private Sub txtAPIKey_GotFocus()
    objTextBox.AutoSelect txtAPIKey
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


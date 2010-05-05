VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmScribdTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scribd. Wrapper Test"
   ClientHeight    =   10005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   Icon            =   "frmScribdTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserId 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   5130
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtSessionKey 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   7650
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2610
      TabIndex        =   2
      Top             =   1110
      Width           =   2325
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   90
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
   Begin VB.TextBox txtDocumentOffset 
      Height          =   315
      Left            =   7830
      TabIndex        =   12
      Top             =   4500
      Width           =   2325
   End
   Begin VB.TextBox txtDocumentLimit 
      Height          =   315
      Left            =   5250
      TabIndex        =   11
      Top             =   4500
      Width           =   2325
   End
   Begin VB.CheckBox chkPrivate 
      Caption         =   "&Private"
      Height          =   405
      Left            =   90
      TabIndex        =   8
      Top             =   3420
      Width           =   1845
   End
   Begin VB.TextBox txtDocumentId 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4500
      Width           =   2325
   End
   Begin VB.TextBox txtDocumentKey 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4500
      Width           =   2325
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9780
      Top             =   5340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   345
      Left            =   7590
      TabIndex        =   7
      Top             =   3030
      Width           =   825
   End
   Begin VB.ComboBox cboMethod 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1980
      Width           =   4905
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   525
      Left            =   1380
      TabIndex        =   14
      Top             =   5310
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   13
      Top             =   5310
      Width           =   1245
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   90
      TabIndex        =   6
      Top             =   3030
      Width           =   7455
   End
   Begin VB.TextBox txtInputBuffer 
      Height          =   3195
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   6600
      Width           =   5145
   End
   Begin VB.TextBox txtOutputBuffer 
      Height          =   3195
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   6600
      Width           =   5025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   9
      X1              =   -30
      X2              =   10970
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   8
      X1              =   -30
      X2              =   10970
      Y1              =   5085
      Y2              =   5085
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   7
      X1              =   0
      X2              =   11000
      Y1              =   4035
      Y2              =   4035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   0
      X2              =   11000
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   5
      X1              =   0
      X2              =   11000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   4
      X1              =   0
      X2              =   11000
      Y1              =   1575
      Y2              =   1575
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
      Left            =   5130
      TabIndex        =   29
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
      Left            =   7650
      TabIndex        =   28
      Top             =   870
      Width           =   1050
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
      Left            =   2610
      TabIndex        =   27
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
      Left            =   90
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   90
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Document Offset"
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
      Index           =   12
      Left            =   7800
      TabIndex        =   24
      Top             =   4260
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Document Limit"
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
      Index           =   11
      Left            =   5220
      TabIndex        =   23
      Top             =   4260
      Width           =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   11000
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   0
      X2              =   11000
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Document Id"
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
      Index           =   9
      Left            =   90
      TabIndex        =   22
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Document Key"
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
      Index           =   8
      Left            =   2610
      TabIndex        =   21
      Top             =   4260
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   11000
      Y1              =   6015
      Y2              =   6015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   11000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Method"
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
      Index           =   6
      Left            =   90
      TabIndex        =   20
      Top             =   1740
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
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
      TabIndex        =   19
      Top             =   2790
      Width           =   855
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
      Left            =   5190
      TabIndex        =   18
      Top             =   6360
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
      Left            =   90
      TabIndex        =   17
      Top             =   6360
      Width           =   2100
   End
End
Attribute VB_Name = "frmScribdTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- A basic application which allows the user to test Scribd API functions including uploading of
'--- documents.
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
Private objLastValue                    As New LastValue
Private objLastPosition                 As New LastPosition
Private objTextBox                      As New TextBoxClass
Private objFileManagement               As New FileManagement

'--- Happens when the user clicks on the Browse button. They will be shown a dialog
'--- where they can select the file they want to upload.
Private Sub cmdBrowse_Click()

    With CommonDialog1
    
        .ShowOpen
        
        If Not .CancelError Then
            txtFileName = .FileName
        End If
        
    End With
    
End Sub

'--- When the user clicks on the Ok button then we will attempt to perform the selected
'--- action.
Private Sub cmdOk_Click()

    cmdOk.Enabled = False
    cmdExit.Enabled = False
    
    txtOutputBuffer.Text = ""
    txtInputBuffer.Text = ""

    With objScribd
    
        .APIKey = txtAPIKey
        .Username = txtUsername
        .Password = txtPassword
        .SessionKey = txtSessionKey
        
        .Method = cboMethod.List(cboMethod.ListIndex)
        
        Select Case cboMethod.List(cboMethod.ListIndex)
        Case "docs.getList"
            
            If Not Len(txtDocumentLimit) = 0 Then
                .DocumentLimit = CInt(txtDocumentLimit)
            End If
            If Not Len(txtDocumentOffset) = 0 Then
                .DocumentOffset = CInt(txtDocumentOffset)
            End If
        
        Case "docs.delete"
            
            .CreateDocument
            .Document.DocumentId = txtDocumentId
        
        Case "docs.upload"
            
            .CreateDocument
            .MimeType = "application/octet-stream"
            .Document.FileName = txtFileName
            .Document.DocumentType = objFileManagement.GetFileType(txtFileName)
            If chkPrivate.Value = 1 Then
                .Document.Access = scribdPrivateDocument
            Else
                .Document.Access = scribdPublicDocument
            End If
        
        End Select
        
        If .DoAction Then
        
            Select Case cboMethod.List(cboMethod.ListIndex)
            Case "docs.getList"
                
                If Not .DocumentCount = 0 Then
                    .DocumentNumber = 0
                    MsgBox .DocumentCount & " documents were found. The first document title was '" & .Document.Title & "'."
                Else
                    MsgBox "No documents were found."
                End If
                
            Case "docs.delete"
            Case "docs.upload"
                
                txtDocumentId = .Document.DocumentId
                txtDocumentKey = .Document.AccessKey
            
            Case "user.login"
                
                txtUserId = .UserId
                txtSessionKey = .SessionKey
            
            End Select
            
            txtInputBuffer.Text = .Result
        
        Else
            MsgBox "Method '" & cboMethod.List(cboMethod.ListIndex) & "' in the Scribd API failed with the following error: " & .ErrorNumber & " - " & .ErrorDescription, vbOKOnly + vbCritical, "Scribd API Wrapper"
            txtInputBuffer.Text = .Result
        End If
    
    End With
    
    cmdOk.Enabled = True
    cmdExit.Enabled = True
    
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub

'--- This event is raised by the Scribd Object whenever it sends data to Scribd
Private Sub objScribd_PostingData(DataToPost As Variant)
    txtOutputBuffer.Text = txtOutputBuffer.Text & DataToPost
    DoEvents
End Sub

Private Sub Form_Load()

    Set objScribd = New Scribd_VB6
    
    cboMethod.AddItem "docs.getList"
    cboMethod.AddItem "docs.delete"
    cboMethod.AddItem "docs.upload"
    cboMethod.AddItem "user.login"

    With objLastValue
        .GetLastValue App.EXEName, txtAPIKey
        .GetLastValue App.EXEName, cboMethod
        .GetLastValue App.EXEName, txtUsername
        .GetLastValue App.EXEName, txtPassword
        .GetLastValue App.EXEName, txtUserId
        .GetLastValue App.EXEName, txtSessionKey
        .GetLastValue App.EXEName, txtFileName
        .GetLastValue App.EXEName, chkPrivate
        .GetLastValue App.EXEName, txtDocumentId
        .GetLastValue App.EXEName, txtDocumentKey
        .GetLastValue App.EXEName, txtDocumentLimit
        .GetLastValue App.EXEName, txtDocumentOffset
        .GetLastValue App.EXEName, txtOutputBuffer
        .GetLastValue App.EXEName, txtInputBuffer
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set objScribd = Nothing
    
    With objLastValue
        .SaveLastValue App.EXEName, txtAPIKey
        .SaveLastValue App.EXEName, cboMethod
        .SaveLastValue App.EXEName, txtUsername
        .SaveLastValue App.EXEName, txtPassword
        .SaveLastValue App.EXEName, txtUserId
        .SaveLastValue App.EXEName, txtSessionKey
        .SaveLastValue App.EXEName, txtFileName
        .SaveLastValue App.EXEName, chkPrivate
        .SaveLastValue App.EXEName, txtDocumentId
        .SaveLastValue App.EXEName, txtDocumentKey
        .SaveLastValue App.EXEName, txtDocumentLimit
        .SaveLastValue App.EXEName, txtDocumentOffset
        .SaveLastValue App.EXEName, txtOutputBuffer
        .SaveLastValue App.EXEName, txtInputBuffer
    End With

End Sub

'--- These events happen when the textbox controls gain focus
Private Sub txtAPIKey_GotFocus()
    objTextBox.AutoSelect txtAPIKey
End Sub

Private Sub txtDocumentId_GotFocus()
    objTextBox.AutoSelect txtDocumentId
End Sub

Private Sub txtDocumentKey_GotFocus()
    objTextBox.AutoSelect txtDocumentKey
End Sub

Private Sub txtDocumentLimit_GotFocus()
    objTextBox.AutoSelect txtDocumentLimit
End Sub

Private Sub txtDocumentLimit_KeyPress(KeyAscii As Integer)
    objTextBox.ValidateKeyPress txtDocumentLimit, tbsNumeric, KeyAscii
End Sub

Private Sub txtDocumentOffset_GotFocus()
    objTextBox.AutoSelect txtDocumentOffset
End Sub

Private Sub txtDocumentOffset_KeyPress(KeyAscii As Integer)
    objTextBox.ValidateKeyPress txtDocumentOffset, tbsNumeric, KeyAscii
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


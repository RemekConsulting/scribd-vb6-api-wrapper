VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scribd_VB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- Implements a wrapper for the Scribd API to be used with VB6 (The forgotten programming language).
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

'--- XML declarations
Private objXML                          As MSXML2.DOMDocument
Private objElement                      As MSXML2.IXMLDOMElement
Private objNodeList                     As MSXML2.IXMLDOMNodeList
Private objNode                         As MSXML2.IXMLDOMNode

'--- Configuration properties are stored here
Private propAPI_URL                     As String
Private propAPI_Script                  As String
Private propAPIKey                      As String
Private propUsername                    As String
Private propPassword                    As String
Private propSessionKey                  As String
Private propUserId                      As Long

'--- Action properties are stored here
Private propMethod                      As String
Private propResult                      As String
Private propMimeType                    As String
Private propContentType                 As String
Private propDocumentLimit               As Integer
Private propDocumentOffset              As Integer

'---Document attributes are stored in a separate class. That class is accessed through these properties
Private propDocuments()                 As Scribd_Document
Private propDocumentCount               As Long
Private propDocumentNumber              As Integer

'--- Error handling variables
Private propErrorNumber                 As String
Private propErrorDescription            As String

'--- Working variables are declared here
Private strPostData                     As String
Private strBoundary                     As String
Private lngPointer                      As Long

'--- HTTP activity declarations
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_SERVICE_HTTP = 3
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Private Const scUserAgent = "Scribd VB6 Wrapper"

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszCallerName As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal lpszServerName As String, ByVal nProxyPort As Integer, ByVal lpszUsername As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hInternetSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, ByVal lpszReferer As String, ByVal lpszAcceptTypes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Boolean
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInternetHandle As Long) As Boolean

'--- Event raised by the Class when posting data to the web server
Event PostingData(DataToPost)

'-------------------------------------------------------------------'
'--- P U B L I C   F U N C T I O N S   D E C L A R E D   H E R E ---'
'-------------------------------------------------------------------'
Public Function CreateDocument() As Boolean

    propDocumentCount = 1
    propDocumentNumber = 0
    ReDim propDocuments(propDocumentNumber)
    Set propDocuments(propDocumentNumber) = New Scribd_Document

End Function

Public Function DoAction() As Boolean

    Dim boolResult                      As Boolean
    
    Dim lngTotalDocuments               As Long
    
    On Error GoTo ErrorHandler
    
    '---Set the default Content Type
    propContentType = "Content-Type: application/x-www-form-urlencoded"

    '--- Assemble the standard post data
    strPostData = "method=" & propMethod & "&api_key=" & propAPIKey

    If Not Len(propSessionKey) = 0 Then
        strPostData = strPostData & "&session_key=" & propSessionKey
    ElseIf Not propUserId = 0 Then
        strPostData = strPostData & "&my_user_id=" & CStr(propUserId)
    End If
    
    '--- Now assemble the remaining post data based on the method
    Select Case propMethod
    Case "user.login"
        
        strPostData = strPostData & "&username=" & propUsername
        strPostData = strPostData & "&password=" & propPassword
        
        propUserId = 0
        propSessionKey = ""
    
    Case "docs.changeSettings"
        
        With propDocuments(propDocumentNumber)
        
            strPostData = strPostData & "&doc_ids=" & CStr(.DocumentId)
            
            If Not Len(.ISBN) = 0 Then
                strPostData = strPostData & "&isbn=" & .ISBN
            End If
            If Not Len(.Title) = 0 Then
                strPostData = strPostData & "&title=" & .Title
            End If
            If Not Len(.Description) = 0 Then
                strPostData = strPostData & "&description=" & .Description
            End If
            If Not Len(.AccessString) = 0 Then
                strPostData = strPostData & "&access=" & .AccessString
            End If
            If Not Len(.LicenseString) = 0 Then
                strPostData = strPostData & "&license=" & .LicenseString
            End If
            If Not Len(.ShowAdverts) = 0 Then
                strPostData = strPostData & "&show_ads=" & .ShowAdverts
            End If
            If Not Len(.LinkBackURL) = 0 Then
                strPostData = strPostData & "&link_back_url=" & .LinkBackURL
            End If
            If Not Len(.Category) = 0 Then
                strPostData = strPostData & "&category=" & .Category
            End If
            If Not Len(.Tags) = 0 Then
                strPostData = strPostData & "&tags=" & .Tags
            End If
            strPostData = strPostData & "&download_formats=" & .DownloadFormatsString
            If Not Len(.Author) = 0 Then
                strPostData = strPostData & "&author=" & .Author
            End If
            If Not Len(.Publisher) = 0 Then
                strPostData = strPostData & "&publisher=" & .Publisher
            End If
            If Not Len(.WhenPublished) = 0 Then
                strPostData = strPostData & "&when_published=" & .WhenPublished
            End If
            If Not Len(.Edition) = 0 Then
                strPostData = strPostData & "&edition=" & .Edition
            End If
            
            
            If .DisableUploadLink Then
                strPostData = strPostData & "&disable_upload_link=1"
            Else
                strPostData = strPostData & "&disable_upload_link=0"
            End If
            
            If .DisablePrint Then
                strPostData = strPostData & "&disable_print=1"
            Else
                strPostData = strPostData & "&disable_print=0"
            End If
            
            If .DisableSelectText Then
                strPostData = strPostData & "&disable_select_text=1"
            Else
                strPostData = strPostData & "&disable_select_text=0"
            End If
            
            If .DisableAboutDialog Then
                strPostData = strPostData & "&disable_about_dialog=1"
            Else
                strPostData = strPostData & "&disable_about_dialog=0"
            End If
            
            If .DisableInfoDialog Then
                strPostData = strPostData & "&disable_info_dialog=1"
            Else
                strPostData = strPostData & "&disable_info_dialog=0"
            End If
            
            If .DisableViewModeChange Then
                strPostData = strPostData & "&disable_view_mode_change=1"
            Else
                strPostData = strPostData & "&disable_view_mode_change=0"
            End If
            
            If .DisableRelatedDocs Then
                strPostData = strPostData & "&disable_related_docs=1"
            Else
                strPostData = strPostData & "&disable_related_docs=0"
            End If
            
        End With
    
    Case "docs.delete"
        
        strPostData = strPostData & "&doc_id=" & CStr(propDocuments(propDocumentNumber).DocumentId)
    
    Case "docs.upload"
        
        PreFileUpload
    
    Case "docs.getList"
        
        If Not propDocumentLimit = 0 Then
            strPostData = strPostData & "&limit=" & CStr(propDocumentLimit)
        End If
        
        If Not propDocumentOffset = 0 Then
            strPostData = strPostData & "&offset=" & CStr(propDocumentOffset)
        End If
        
        Erase propDocuments
        propDocumentCount = 0
    
    End Select

    boolResult = PostData(propAPI_URL, propAPI_Script, strPostData, propContentType)

    '--- Now process the result
    ProcessResponse propResult
    
    Set objElement = objXML.selectSingleNode("//rsp")
    If Not objElement Is Nothing Then
    
    If objElement.getAttribute("stat") = "ok" Then
    
        Select Case propMethod
        Case "user.login"
            propSessionKey = objXML.selectSingleNode("//session_key").Text
            propUserId = CLng(objXML.selectSingleNode("//user_id").Text)
        Case "docs.delete"
        Case "docs.upload"
            propDocuments(propDocumentNumber).DocumentId = CLng(objXML.selectSingleNode("//doc_id").Text)
            propDocuments(propDocumentNumber).AccessKey = objXML.selectSingleNode("//access_key").Text
        Case "docs.getList"
        
            Set objElement = objXML.selectSingleNode("//resultset")
            
            lngTotalDocuments = objElement.getAttribute("totalResultsReturned")
            ReDim propDocuments(lngTotalDocuments) As Scribd_Document
            lngPointer = 0
            
            Set objNodeList = objXML.getElementsByTagName("result")
            
            For Each objNode In objNodeList
            
                Set propDocuments(lngPointer) = New Scribd_Document
                
                With propDocuments(lngPointer)
                    .DocumentId = CLng(objNode.selectSingleNode("doc_id").Text)
                    .AccessKey = objNode.selectSingleNode("access_key").Text
                    .SecretPassword = objNode.selectSingleNode("secret_password").Text
                    .Title = objNode.selectSingleNode("title").Text
                    .Description = objNode.selectSingleNode("description").Text
                    .ThumbnailURL = objNode.selectSingleNode("thumbnail_url").Text
                    .ConversionStatus = objNode.selectSingleNode("conversion_status").Text
                    If Not objNode.selectSingleNode("page_count").Text = "" Then
                        .PageCount = CInt(objNode.selectSingleNode("page_count").Text)
                    End If
                End With
                
                lngPointer = lngPointer + 1
                
                propDocumentCount = lngPointer
                
            Next objNode
            
        End Select
    
        DoAction = True
        
    ElseIf objElement.getAttribute("stat") = "fail" Then
    
        Set objElement = objXML.selectSingleNode("//error")
        
        propErrorNumber = objElement.getAttribute("code")
        propErrorDescription = objElement.getAttribute("message")
    
    End If
    
    End If
    
ErrorHandler:

    Select Case Err.Number
    Case 0
    Case Else
        propErrorNumber = Err.Number
        propErrorDescription = Err.Description
    End Select
    
    Exit Function
    Resume

End Function


'---------------------------------------------------------------------'
'--- P R I V A T E   F U N C T I O N S   D E C L A R E D   H E R E ---'
'---------------------------------------------------------------------'
Private Function ProcessResponse(ByVal ResponseText As String) As Boolean

    Set objXML = New MSXML2.DOMDocument
    
    If objXML.loadXML(ResponseText) Then
        propErrorNumber = objXML.parseError.errorCode
        propErrorDescription = objXML.parseError.reason
    End If

End Function

'--- This function prepares the MIME boundaries and parts for a Multipart posting
'--- to Scribd to upload a document.
Private Function PreFileUpload() As Boolean

    Dim strFileContents                 As String
    
    On Error GoTo ErrorHandler
    
    '--- Create a MIME boundary consisting of a random string
    strBoundary = MimeBoundary(32)
    
    '--- Set the content type
    propContentType = "Content-Type: multipart/form-data; boundary=" & strBoundary
    
    strPostData = "--" & strBoundary & vbCrLf
    strPostData = strPostData & "Content-Disposition: form-data; name=""method""" & vbCrLf
    strPostData = strPostData & vbCrLf & "docs.upload" & vbCrLf
    
    strPostData = strPostData & "--" & strBoundary & vbCrLf
    strPostData = strPostData & "Content-Disposition: form-data; name=""api_key""" & vbCrLf
    strPostData = strPostData & vbCrLf & propAPIKey & vbCrLf

    With propDocuments(propDocumentNumber)
    
        If Not Len(.DocumentType) = 0 Then
            strPostData = strPostData & "--" & strBoundary & vbCrLf
            strPostData = strPostData & "Content-Disposition: form-data; name=""doctype""" & vbCrLf
            strPostData = strPostData & vbCrLf & .DocumentType & vbCrLf
        End If
    
        If Not Len(.DocumentRevision) = 0 Then
            strPostData = strPostData & "--" & strBoundary & vbCrLf
            strPostData = strPostData & "Content-Disposition: form-data; name=""rev_id""" & vbCrLf
            strPostData = strPostData & vbCrLf & .DocumentRevision & vbCrLf
        End If
    
        strPostData = strPostData & "--" & strBoundary & vbCrLf
        strPostData = strPostData & "Content-Disposition: form-data; name=""access""" & vbCrLf
        strPostData = strPostData & vbCrLf & .AccessString & vbCrLf
    
        If Not Len(propSessionKey) = 0 Then
            strPostData = strPostData & "--" & strBoundary & vbCrLf
            strPostData = strPostData & "Content-Disposition: form-data; name=""session_key""" & vbCrLf
            strPostData = strPostData & vbCrLf & propSessionKey & vbCrLf
        ElseIf Not propUserId = 0 Then
            strPostData = strPostData & "--" & strBoundary & vbCrLf
            strPostData = strPostData & "Content-Disposition: form-data; name=""my_user_id""" & vbCrLf
            strPostData = strPostData & vbCrLf & CStr(propUserId) & vbCrLf
        End If
        
        '--- Read the file contents as a string
        '--- NOTE: in HTTP everything is a string, even binary files
        strFileContents = GetFileContents(.FileName)
        
        strPostData = strPostData & "--" & strBoundary & vbCrLf
        strPostData = strPostData & "Content-Disposition: form-data; name=""file""; filename=""" & GetFileName(.FileName) & """" & vbCrLf
        strPostData = strPostData & "Content-Type: " & propMimeType & vbCrLf
        strPostData = strPostData & vbCrLf & strFileContents
        
        strPostData = strPostData & vbCrLf & "--" & strBoundary & "--"

    End With
    
ErrorHandler:

    Select Case Err.Number
    Case 0
    Case Else
    End Select
    
    Exit Function
    Resume

End Function

'--- This function generates a MIME boundary string of a given length. It is used to
'--- create MIME boundaries in a Multi-Part request.
Private Function MimeBoundary(ByVal BoundaryLength As Integer)
    
    Dim strWorker                       As String
    
    Dim intPointer                      As Integer
    Dim byteWorker                      As Byte
    
    Randomize
    
    For intPointer = 1 To BoundaryLength
        
        byteWorker = Int(Rnd() * 127)
    
        If (byteWorker >= Asc("0") And byteWorker <= Asc("9")) Or (byteWorker >= Asc("A") And byteWorker <= Asc("Z")) Or (byteWorker >= Asc("a") And byteWorker <= Asc("z")) Then
            strWorker = strWorker & Chr(byteWorker)
        Else
            intPointer = intPointer - 1
        End If
    
    Next intPointer
    
    MimeBoundary = strWorker
    
End Function


'--- This function retireves the contents of a file and returns it as a string
'--- this is also ture for binary files
Private Function GetFileContents(ByVal FileName As String) As String
    
    Dim strWorker                       As String
    Dim lngLength                       As Long
    
    lngLength = FileLen(FileName)
    strWorker = String(lngLength, Chr(0))
    
    On Error GoTo ErrorHandler
    
    Open FileName For Binary As #1
    
    Get #1, , strWorker
    
    GetFileContents = strWorker
    
    Close #1
    
ErrorHandler:

    Select Case Err.Number
    Case 0
    Case Else
        MsgBox Err.Description, vbCritical, "ERROR"
    End Select
    
    Err.Clear
    Exit Function
    Resume
    
End Function

Private Function GetFileName(CheckFile) As String

    On Error GoTo GetFileName_Error
    
    Dim strFileName             As String
    
    GetFileName = CheckFile
    
    If Not InStrRev(CheckFile, "\") = 0 Then GetFileName = Mid$(CheckFile, InStrRev(CheckFile, "\") + 1)
    
GetFileName_Error:

    Select Case Err.Number
    Case 0
    Case Else
        MsgBox Err.Description, vbCritical, "ERROR"
    End Select
    
    Exit Function
    Resume
    
End Function

'--- The following function is used to send post data over http to the Scribd API server. Any
'--- returned XML is loaded into the Result property which can then be decrypted by the calling
'--- software.
Private Function PostData(Server As String, Script As String, DataToPost As String, ContentType As String) As String
    
    Dim hInternetOpen                   As Long
    Dim hInternetConnect                As Long
    Dim hHttpOpenRequest                As Long
    Dim bRet                            As Boolean
    Dim bDoLoop                         As Boolean
    Dim sReadBuffer                     As String * 4096
    Dim lNumberOfBytesRead              As Long
    Dim sBuffer                         As String
    Dim sHeader                         As String
    Dim lPostDataLen                    As Long
    
    PostData = False
    
    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0
    
    '-- Clear any previous result
    propResult = ""
    
    hInternetOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    If hInternetOpen <> 0 Then
    
        '--- Change the server to your server name
        hInternetConnect = InternetConnect(hInternetOpen, Server, INTERNET_DEFAULT_HTTP_PORT, vbNullString, "HTTP/1.0", INTERNET_SERVICE_HTTP, 0, 0)
        
        If hInternetConnect <> 0 Then
            
            '--- Brings the data across the wire even if it locally cached.
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", Script, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
            
            If hHttpOpenRequest <> 0 Then
                
                RaiseEvent PostingData(ContentType & vbCrLf)
                RaiseEvent PostingData("Content-Length: " & Len(DataToPost) & vbCrLf)
                RaiseEvent PostingData(DataToPost)
                lPostDataLen = Len(DataToPost)
                bRet = HttpSendRequest(hHttpOpenRequest, ContentType, Len(ContentType), DataToPost, lPostDataLen)
                
                bDoLoop = True
                Do While bDoLoop
                    sReadBuffer = vbNullString
                    bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                    sBuffer = sBuffer & Left(sReadBuffer, lNumberOfBytesRead)
                    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                Loop
                
                propResult = Replace(sBuffer, vbCrLf, vbLf)
                propResult = Replace(propResult, vbLf, vbCrLf)
                
                bRet = InternetCloseHandle(hHttpOpenRequest)
                PostData = True
                
            End If
            
            bRet = InternetCloseHandle(hInternetConnect)
        
        End If
        
        bRet = InternetCloseHandle(hInternetOpen)
    
    End If

End Function


'---------------------------------------------------------------'
'--- P R O P E R T Y   V A L U E S   A R E   S E T   H E R E ---'
'---------------------------------------------------------------'
Public Property Let API_URL(vData As String)
    propAPI_URL = vData
End Property
Public Property Let API_Script(vData As String)
    propAPI_Script = vData
End Property
Public Property Let APIKey(vData As String)
    propAPIKey = vData
End Property
Public Property Let Username(vData As String)
    propUsername = vData
End Property
Public Property Let Password(vData As String)
    propPassword = vData
End Property
Public Property Let SessionKey(vData As String)
    propSessionKey = vData
End Property
Public Property Let UserId(vData As Long)
    propUserId = vData
End Property
Public Property Let Method(vData As String)
    propMethod = vData
End Property
Public Property Let DocumentLimit(vData As Integer)
    propDocumentLimit = vData
End Property
Public Property Let DocumentOffset(vData As Integer)
    propDocumentOffset = vData
End Property
Public Property Let MimeType(vData As String)
    propMimeType = vData
End Property
Public Property Let DocumentNumber(vData As Integer)
    propDocumentNumber = vData
End Property


'---------------------------------------------------------------------------'
'--- P R O P E R T Y   V A L U E S   A R E   R E T R I E V E D   H E R E ---'
'---------------------------------------------------------------------------'
Public Property Get Result() As String
    Result = propResult
End Property
Public Property Get SessionKey() As String
    SessionKey = propSessionKey
End Property
Public Property Get UserId() As Long
    UserId = propUserId
End Property
Public Property Get DocumentCount() As Long
    DocumentCount = propDocumentCount
End Property
Public Property Get Document() As Scribd_Document
    Set Document = propDocuments(propDocumentNumber)
End Property
Public Property Get DocumentNumber() As Integer
    DocumentNumber = propDocumentNumber
End Property
Public Property Get ErrorNumber() As String
    ErrorNumber = propErrorNumber
End Property
Public Property Get ErrorDescription() As String
    ErrorDescription = propErrorDescription
End Property


'---------------------------------------------------------------------------------------------'
'--- C L A S S   C R E A T I O N   A N D   T E R M I N A T I O N   H A N D L E D   H E R E ---'
'---------------------------------------------------------------------------------------------'
Private Sub Class_Initialize()

    propAPI_URL = "api.scribd.com"
    propAPI_Script = "/api?"
    propMimeType = "application/octet-stream"
    
    CreateDocument

End Sub

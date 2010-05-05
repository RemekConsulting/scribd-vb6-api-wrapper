VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Scribd_Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- Implements a wrapper for the Scribd API to be used with VB6 (The forgotten programming language).
'--- This class implements an instance of a document which is used by the Scribd_VB6 base class. We use
'--- an external class so that the calling software can also work with the Scribd Document class.
'---
'--- AUTHOR:        Greg Bridle
'--- DATE:          2010.04.16
'---
'--- DEPENDANCIES:  Scribd_VB6.cls (Base class)
'---                Microsoft XML 6.0 (Used by the Base class)
'---
'--- PATCH HISTORY
'---
'--- DATE           BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

'--- Enumerated property values are declared here. We use Enums to reduce the possibility of errors
'--- caused by the wrong value being supplied to the Scribd API. At the end of this class we also
'--- provide Friend properties which can be used inter-class to directly access the string value.
Public Enum DocumentAccess
    scribdPrivateDocument = 0
    scribdPublicDocument = 1
End Enum

Public Enum License
    scribdLicense_Copyright = 0
    scribdLicense_PublicDomain = 1
    scribdLicense_Attribution = 2
    scribdLicense_Attribution_NonCommercial = 3
    scribdLicense_Attribution_NonCommercial_NoDerivs = 4
    scribdLicense_Attribution_NonCommercial_ShareAlike = 5
    scribdLicense_Attribution_ShareAlike = 6
    scribdLicense_Attribution_NoDerivs = 7
End Enum

Public Enum DownloadFormat
    scribdDownload_None = 0
    scribdDownload_PDF = 1
    scribdDownload_Text = 2
    scribdDownload_Original = 3
End Enum

'--- All other document property values are declared here.
Private propDocumentId                  As Long
Private propDocumentType                As String
Private propDocumentRevision            As String
Private propFileName                    As String
Private propAccessKey                   As String
Private propSecretPassword              As String
Private propConversionStatus            As String
Private propPageCount                   As Integer
Private propThumbnailURL                As String
Private propISBN                        As String
Private propTitle                       As String
Private propDescription                 As String
Private propAccess                      As String
Private propLicense                     As String
Private propShowAdverts                 As String
Private propLinkBackURL                 As String
Private propCategory                    As String
Private propTags                        As String
Private propDownloadFormats             As String
Private propAuthor                      As String
Private propPublisher                   As String
Private propWhenPublished               As String
Private propEdition                     As String
Private propDisableUploadLink           As Boolean
Private propDisablePrint                As Boolean
Private propDisableSelectText           As Boolean
Private propDisableAboutDialog          As Boolean
Private propDisableInfoDialog           As Boolean
Private propDisableViewModeChange       As Boolean
Private propDisableRelatedDocs          As Boolean

'---------------------------------------------------------------'
'--- P R O P E R T Y   V A L U E S   A R E   S E T   H E R E ---'
'---------------------------------------------------------------'
Public Property Let DocumentId(vData As Long)
    propDocumentId = vData
End Property
Public Property Let DocumentType(vData As String)
    propDocumentType = vData
End Property
Public Property Let DocumentRevision(vData As String)
    propDocumentRevision = vData
End Property
Public Property Let FileName(vData As String)
    propFileName = vData
End Property
Public Property Let AccessKey(vData As String)
    propAccessKey = vData
End Property
Public Property Let SecretPassword(vData As String)
    propSecretPassword = vData
End Property
Public Property Let ConversionStatus(vData As String)
    propConversionStatus = vData
End Property
Public Property Let PageCount(vData As Integer)
    propPageCount = vData
End Property
Public Property Let ThumbnailURL(vData As String)
    propThumbnailURL = vData
End Property
Public Property Let ISBN(vData As String)
    propISBN = vData
End Property
Public Property Let Title(vData As String)
    propTitle = vData
End Property
Public Property Let Description(vData As String)
    propDescription = vData
End Property
Public Property Let ShowAdverts(vData As String)
    propShowAdverts = vData
End Property
Public Property Let LinkBackURL(vData As String)
    propLinkBackURL = vData
End Property
Public Property Let Category(vData As String)
    propCategory = vData
End Property
Public Property Let Tags(vData As String)
    propTags = vData
End Property
Public Property Let Author(vData As String)
    propAuthor = vData
End Property
Public Property Let Publisher(vData As String)
    propPublisher = vData
End Property
Public Property Let WhenPublished(vData As String)
    propWhenPublished = vData
End Property
Public Property Let Edition(vData As String)
    propEdition = vData
End Property
Public Property Let DisableUploadLink(vData As Boolean)
    propDisableUploadLink = vData
End Property
Public Property Let DisablePrint(vData As Boolean)
    propDisablePrint = vData
End Property
Public Property Let DisableSelectText(vData As Boolean)
    propDisableSelectText = vData
End Property
Public Property Let DisableAboutDialog(vData As Boolean)
    propDisableAboutDialog = vData
End Property
Public Property Let DisableInfoDialog(vData As Boolean)
    propDisableInfoDialog = vData
End Property
Public Property Let DisableViewModeChange(vData As Boolean)
    propDisableViewModeChange = vData
End Property
Public Property Let DisableRelatedDocs(vData As Boolean)
    propDisableRelatedDocs = vData
End Property


'---------------------------------------------------------------------------'
'--- P R O P E R T Y   V A L U E S   A R E   R E T R I E V E D   H E R E ---'
'---------------------------------------------------------------------------'
Public Property Get DocumentId() As Long
    DocumentId = propDocumentId
End Property
Public Property Get DocumentType() As String
    DocumentType = propDocumentType
End Property
Public Property Get DocumentRevision() As String
    DocumentRevision = propDocumentRevision
End Property
Public Property Get FileName() As String
    FileName = propFileName
End Property
Public Property Get AccessKey() As String
    AccessKey = propAccessKey
End Property
Public Property Get SecretPassword() As String
    SecretPassword = propSecretPassword
End Property
Public Property Get ConversionStatus() As String
    ConversionStatus = propConversionStatus
End Property
Public Property Get PageCount() As Integer
    PageCount = propPageCount
End Property
Public Property Get ThumbnailURL() As String
    ThumbnailURL = propThumbnailURL
End Property
Public Property Get ISBN() As String
    ISBN = propISBN
End Property
Public Property Get Title() As String
    Title = propTitle
End Property
Public Property Get Description() As String
    Description = propDescription
End Property
Public Property Get ShowAdverts() As String
    ShowAdverts = propShowAdverts
End Property
Public Property Get LinkBackURL() As String
    LinkBackURL = propLinkBackURL
End Property
Public Property Get Category() As String
    Category = propCategory
End Property
Public Property Get Tags() As String
    Tags = propTags
End Property
Public Property Get Author() As String
    Author = propAuthor
End Property
Public Property Get Publisher() As String
    Publisher = propPublisher
End Property
Public Property Get WhenPublished() As String
    WhenPublished = propWhenPublished
End Property
Public Property Get Edition() As String
    Edition = propEdition
End Property
Public Property Get DisableUploadLink() As Boolean
    DisableUploadLink = propDisableUploadLink
End Property
Public Property Get DisablePrint() As Boolean
    DisablePrint = propDisablePrint
End Property
Public Property Get DisableSelectText() As Boolean
    DisableSelectText = propDisableSelectText
End Property
Public Property Get DisableAboutDialog() As Boolean
    DisableAboutDialog = propDisableAboutDialog
End Property
Public Property Get DisableInfoDialog() As Boolean
    DisableInfoDialog = propDisableInfoDialog
End Property
Public Property Get DisableViewModeChange() As Boolean
    DisableViewModeChange = propDisableViewModeChange
End Property
Public Property Get DisableRelatedDocs() As Boolean
    DisableRelatedDocs = propDisableRelatedDocs
End Property

'-----------------------------------------------------------------------------'
'--- P R O P E R T Y   V A L U E S   W H I C H   A R E   P R O C E S S E D ---'
'-----------------------------------------------------------------------------'
Public Property Let Access(vData As DocumentAccess)
    
    If vData = scribdPrivateDocument Then
        propAccess = "private"
    ElseIf vData = scribdPublicDocument Then
        propAccess = "public"
    Else
        Err.Raise vbObjectError, "Access", "Invalid Access value has been supplied."
    End If

End Property

Public Property Let DownloadFormats(vData As DownloadFormat)
    
    Select Case vData
    Case scribdDownload_None
        propDownloadFormats = ""
    Case scribdDownload_PDF
        propDownloadFormats = "pdf"
    Case scribdDownload_Text
        propDownloadFormats = "text"
    Case scribdDownload_Original
        propDownloadFormats = "original"
    End Select
    
End Property

Public Property Let License(vData As License)

    Select Case vData
    Case scribdLicense_Copyright
        propLicense = "c"
    Case scribdLicense_PublicDomain
        propLicense = "pd"
    Case scribdLicense_Attribution
        propLicense = "by"
    Case scribdLicense_Attribution_NonCommercial
        propLicense = "by-nc"
    Case scribdLicense_Attribution_NonCommercial_NoDerivs
        propLicense = "by-nc-nd"
    Case scribdLicense_Attribution_NonCommercial_ShareAlike
        propLicense = "by-nc-sa"
    Case scribdLicense_Attribution_ShareAlike
        propLicense = "by-sa"
    Case scribdLicense_Attribution_NoDerivs
        propLicense = "by-nd"
    End Select

End Property

Public Property Get Access() As DocumentAccess
    
    If propAccess = "private" Then
        Access = scribdPrivateDocument
    ElseIf propAccess = "public" Then
        Access = scribdPublicDocument
    End If

End Property

Public Property Get DownloadFormats() As DownloadFormat
    
    Select Case propDownloadFormats
    Case ""
        DownloadFormats = scribdDownload_None
    Case "pdf"
        DownloadFormats = scribdDownload_PDF
    Case "text"
        DownloadFormats = scribdDownload_Text
    Case "original"
        DownloadFormats = scribdDownload_Original
    End Select
    
End Property

Public Property Get License() As License
    
    Select Case propLicense
    Case "c"
        License = scribdLicense_Copyright
    Case "pd"
        License = scribdLicense_PublicDomain
    Case "by"
        License = scribdLicense_Attribution
    Case "by-nc"
        License = scribdLicense_Attribution_NonCommercial
    Case "by-nc-nd"
        License = scribdLicense_Attribution_NonCommercial_NoDerivs
    Case "by-nc-sa"
        License = scribdLicense_Attribution_NonCommercial_ShareAlike
    Case "by-sa"
        License = scribdLicense_Attribution_ShareAlike
    Case "by-nd"
        License = scribdLicense_Attribution_NoDerivs
    End Select

End Property


'---------------------------------------------------------------------'
'--- F R I E N D   P R O P E R T I E S   D E C L A R E D   H E R E ---'
'---------------------------------------------------------------------'

'--- Because we use Enums to make the software more robust, we also declare friend properties
'--- which we can inter-class. This means we don't have to do any messing around with the enums
'--- when we just need the actual string value.

Friend Property Get AccessString() As String
    AccessString = propAccess
End Property
Friend Property Get LicenseString() As String
    LicenseString = propLicense
End Property
Friend Property Get DownloadFormatsString() As String
    DownloadFormatsString = propDownloadFormats
End Property


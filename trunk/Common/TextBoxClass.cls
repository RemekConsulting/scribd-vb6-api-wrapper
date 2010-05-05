VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBoxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- This object is used to provide inline processing of Textbox values.
'---
'--- AUTHOR: Greg Bridle
'--- DATE:   2007.05.24.
'---
'--- PATCH HISTORY
'---
'--- DATE       BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

Public Enum TextboxStyle
    tbsNormalCase = 0
    tbsUpperCase = 1
    tbsLowerCase = 2
    tbsUnsignedNumeric = 3
    tbsNumeric = 4
    tbsDecimal = 5
    tbsCurrency = 6
End Enum

Private propMaximumValue                As Double
Private propMinimumValue                As Double

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type TEXTMETRIC
    tmHeight                            As Long
    tmAscent                            As Long
    tmDescent                           As Long
    tmInternalLeading                   As Long
    tmExternalLeading                   As Long
    tmAveCharWidth                      As Long
    tmMaxCharWidth                      As Long
    tmWeight                            As Long
    tmOverhang                          As Long
    tmDigitizedAspectX                  As Long
    tmDigitizedAspectY                  As Long
    tmFirstChar                         As Byte
    tmLastChar                          As Byte
    tmDefaultChar                       As Byte
    tmBreakChar                         As Byte
    tmItalic                            As Byte
    tmUnderlined                        As Byte
    tmStruckOut                         As Byte
    tmPitchAndFamily                    As Byte
    tmCharSet                           As Byte
End Type

'--- System Local Settings
Private Const LOCALE_SMONDECIMALSEP  As Long = &H16     'Decimal separator
Private Const LOCALE_SMONTHOUSANDSEP As Long = &H17     'Thousand separator
Private Const LOCALE_SCURRENCY As Long = &H14           'Currency sign

Private Declare Function GetThreadLocale Lib "Kernel32" () As Long
Private Declare Function GetUserDefaultLCID Lib "Kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
      
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

Private intDecimalSeparator             As Integer
Private intThousandSeparator            As Integer
Private intCurrencySign                 As Integer

Private strDecimalSeparator             As String
Private strThousandSeparator            As String
Private strCurrencySign                 As String

Private boolSignedNegative              As Boolean
Private boolSignedPositive              As Boolean
Private boolCurrencySigned              As Boolean

Public Sub AutoSelect(Object)

    On Error Resume Next
    
    With Object
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Public Sub ValidateKeyPress(Object, Style As TextboxStyle, ByRef KeyAscii As Integer)

    Dim strPreviousValue                As String
    Dim strCurrentValue                 As String
    
    Select Case Style
    Case tbsNormalCase
    Case tbsUpperCase
        
        KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    Case tbsLowerCase
        
        KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    
    Case Else
        
        strPreviousValue = Object.Text
        
        '--- If there is any selected text then we need to remove it now
        If Not Len(Object.SelText) = 0 Then
            Object.Text = Replace(Object.Text, Object.SelText, "")
        End If
        
        strCurrentValue = Object.Text
        
        '--- Determine if there is any sign in the remaining text
        If Not InStr(strCurrentValue, strCurrencySign) = 0 Then
            boolCurrencySigned = True
        Else
            boolCurrencySigned = False
        End If
        If Not InStr(strCurrentValue, "-") = 0 Then
            boolSignedNegative = True
        Else
            boolSignedNegative = False
        End If
        If Not InStr(strCurrentValue, "+") = 0 Then
            boolSignedPositive = True
        Else
            boolSignedPositive = False
        End If

        '--- Remove the currency sign
        If boolCurrencySigned Then
            strCurrentValue = Replace(strCurrentValue, strCurrencySign, "")
        End If
    
        '--- Remove the sign from any currently signed field
        If boolSignedNegative Then
            strCurrentValue = Replace(strCurrentValue, "-", "")
        End If
        If boolSignedPositive Then
            strCurrentValue = Replace(strCurrentValue, "+", "")
        End If
    
        If (Not KeyAscii < 48 And Not KeyAscii > 57) Then
            strCurrentValue = strCurrentValue & Chr$(KeyAscii)
            KeyAscii = 0
        ElseIf KeyAscii = 8 Then
            boolSignedNegative = False
            If Not Len(Object) > 1 Then
                Object.Text = ""
                KeyAscii = 48
            End If
        ElseIf KeyAscii = 45 And Not Style = tbsUnsignedNumeric Then
            boolSignedNegative = True
            KeyAscii = 0
        ElseIf KeyAscii = 43 And Not Style = tbsUnsignedNumeric Then
            boolSignedPositive = True
            KeyAscii = 0
        ElseIf KeyAscii = intDecimalSeparator And (Style = tbsDecimal Or Style = tbsCurrency) Then
            If Not InStr(strCurrentValue, Chr$(KeyAscii)) > 0 Then
                strCurrentValue = strCurrentValue & Chr$(KeyAscii)
            End If
            KeyAscii = 0
        ElseIf KeyAscii = intCurrencySign And Style = tbsCurrency Then
            boolCurrencySigned = True
            KeyAscii = 0
        Else
            KeyAscii = 0
        End If
    
        '--- Now add back any signs we need to show
        If boolCurrencySigned Then
            strCurrentValue = strCurrencySign & strCurrentValue
        End If
        If boolSignedNegative Then
            strCurrentValue = "-" & strCurrentValue
        End If
        
        '--- If we have a Maximum and/or Minimum value then we check to see that the current value does not exceed
        '--- those limits. If it does then we'll discard the entry.
        If Not Len(strCurrentValue) = 0 And Not strCurrentValue = "." Then
            If Not propMaximumValue = 0 And CDbl(strCurrentValue) > propMaximumValue Then
                strCurrentValue = strPreviousValue
            ElseIf Not propMinimumValue = 0 And CDbl(strCurrentValue) < propMinimumValue Then
                strCurrentValue = strPreviousValue
            End If
        End If
        
        If boolSignedPositive Then
            strCurrentValue = "+" & strCurrentValue
        End If
        
        '--- Now set the text to the new value
        Object.Text = strCurrentValue
        Object.SelStart = Len(strCurrentValue)
        
    End Select

End Sub

'--- This function returns the number of visible lines in a multi line Textbox
Public Function VisibleLines(Object) As Integer

    Dim mRect                           As RECT
    Dim tm                              As TEXTMETRIC
    
    Dim intLines                        As Integer
    
    Dim lngFont                         As Long
    Dim lngOldFont                      As Long
    Dim lngHdc                          As Long
    Dim lngDi                           As Long
    
    Const EM_GETRECT = &HB2
    Const WM_GETFONT = &H31

    '--- Calculate the number of pixels this textbox control uses
    SendMessage Object.hWnd, EM_GETRECT, 0, mRect
    intLines = (mRect.Bottom - mRect.Top)
    
    '--- Now calculate the size of the font in pixels
    lngFont = SendMessage(Object.hWnd, WM_GETFONT, 0, 0&)

    '--- Get a device context to the text control.
    lngHdc = GetDC(Object.hWnd)

    '--- Select in the logical font to obtain the exact font metrics.
    If lngFont <> 0 Then lngOldFont = SelectObject(lngHdc, lngFont)

    lngDi = GetTextMetrics(lngHdc, tm)

    ' Select out the logical font
    If lngFont <> 0 Then lngFont = SelectObject(lngHdc, lngOldFont)

    ' The lines depends on the formatting rectangle and font height
    VisibleLines = (intLines / tm.tmHeight)

    ' Release the device context when done.
    lngDi = ReleaseDC(Object.hWnd, lngHdc)
    
End Function

Public Function AutoSize(Object, Container, Optional UseScaleMode As Boolean = True)

    If Not UseScaleMode Then
        With Object
            .Move .Left, .Top, Container.Width - (.Left * 2), Container.Height - .Top - .Left
        End With
    Else
        With Object
            .Move .Left, .Top, Container.ScaleWidth - (.Left * 2), Container.ScaleHeight - .Top - .Left
        End With
    End If
    
End Function


'-------------------------------------------------
'--- P R I V A T E   F U N C T I O N S
'-------------------------------------------------
Private Function GetCurrentLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

    Dim strReturn                       As String
    Dim lngResult                       As Long

    '--- Call the function passing the Locale type variable to retrieve the required size of
    '--- the string buffer needed
    lngResult = GetLocaleInfo(dwLocaleID, dwLCType, strReturn, Len(strReturn))
    
    '--- if successful..
    If lngResult Then
    
        '--- pad the buffer with spaces
        strReturn = Space$(lngResult)
       
        '--- and call again passing the buffer
        lngResult = GetLocaleInfo(dwLocaleID, dwLCType, strReturn, Len(strReturn))
     
        '--- if successful (lngResult > 0)
        If lngResult Then
      
            '--- lngResult holds the size of the string including the terminating null
            GetCurrentLocaleInfo = Left$(strReturn, lngResult - 1)
      
        End If
   
    End If
    
End Function

Public Property Let MaximumValue(vData As Double)
    propMaximumValue = vData
End Property
Public Property Let MinimumValue(vData As Double)
    propMinimumValue = vData
End Property

'-------------------------------------------------
'--- C L A S S   I N I T I A L I Z A T I O N
'-------------------------------------------------
Private Sub Class_Initialize()

    '--- Get decimal place separator from system settings
    strDecimalSeparator = GetCurrentLocaleInfo(GetUserDefaultLCID(), LOCALE_SMONDECIMALSEP)
    intDecimalSeparator = Asc(strDecimalSeparator)

    '--- Get decimal place separator from system settings
    strThousandSeparator = GetCurrentLocaleInfo(GetUserDefaultLCID(), LOCALE_SMONTHOUSANDSEP)
    intThousandSeparator = Asc(strThousandSeparator)

    '--- Get decimal place separator from system settings
    strCurrencySign = GetCurrentLocaleInfo(GetUserDefaultLCID(), LOCALE_SCURRENCY)
    intCurrencySign = Asc(strCurrencySign)

End Sub

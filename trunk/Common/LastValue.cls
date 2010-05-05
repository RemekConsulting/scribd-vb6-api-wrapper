VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LastValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- This object is used to provide quick saving and reloading of the values held in form controls.
'---
'--- AUTHOR: Greg Bridle
'--- DATE:   2007.05.24.
'---
'--- PATCH HISTORY
'---
'--- DATE       BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

Public Sub SaveLastValue(AppName As String, ByRef FormControl, Optional Property As String = "")

    Dim strObjectName                   As String
    
    On Error Resume Next
    
    strObjectName = FormControl.Name
    strObjectName = strObjectName & "(" & FormControl.Index & ")"
    
    If TypeOf FormControl Is TextBox Then
        SaveSetting AppName, "Values", strObjectName, FormControl.Text
    ElseIf TypeOf FormControl Is CheckBox Then
        SaveSetting AppName, "Values", strObjectName, FormControl.Value
    ElseIf TypeOf FormControl Is ComboBox Then
        SaveSetting AppName, "Values", strObjectName, FormControl.ListIndex
    ElseIf TypeOf FormControl Is OptionButton Then
        SaveSetting AppName, "Values", strObjectName, FormControl.Value
    Else
        Select Case LCase$(Property)
        Case "windowstate"
            SaveSetting AppName, "Values", strObjectName, FormControl.WindowState
        End Select
    End If
    
    Err.Clear
    
End Sub

Public Sub GetLastValue(AppName As String, ByRef FormControl, Optional Property As String = "")

    Dim strObjectName                   As String
    
    On Error Resume Next
    
    strObjectName = FormControl.Name
    strObjectName = strObjectName & "(" & FormControl.Index & ")"
    
    If TypeOf FormControl Is TextBox Then
        FormControl.Text = GetSetting(AppName, "Values", strObjectName, FormControl.Text)
    ElseIf TypeOf FormControl Is CheckBox Then
        FormControl.Value = GetSetting(AppName, "Values", strObjectName, FormControl.Value)
    ElseIf TypeOf FormControl Is ComboBox Then
        If Not Len(Property) = 0 And IsNumeric(Property) Then
            FormControl.ListIndex = GetSetting(AppName, "Values", strObjectName, CInt(Property))
        Else
            FormControl.ListIndex = GetSetting(AppName, "Values", strObjectName, FormControl.ListIndex)
        End If
    ElseIf TypeOf FormControl Is OptionButton Then
        FormControl.Value = GetSetting(AppName, "Values", strObjectName, FormControl.Value)
    Else
        Select Case LCase$(Property)
        Case "windowstate"
            FormControl.WindowState = GetSetting(AppName, "Values", strObjectName, FormControl.WindowState)
        End Select
    End If

    Err.Clear
    
End Sub

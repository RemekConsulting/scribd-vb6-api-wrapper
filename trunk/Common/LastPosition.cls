VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LastPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
'---                            O B J E C T   D E S C R I P T I O N                                 ---'
'------------------------------------------------------------------------------------------------------'
'--- This object is used to provide quick saving and re-positioning of objects on the screen.
'---
'--- AUTHOR: Greg Bridle
'--- DATE:   2007.05.24.
'---
'--- PATCH HISTORY
'---
'--- DATE       BY          DESCRIPTION
'------------------------------------------------------------------------------------------------------'

Option Explicit

Private propSaveWindowSize              As Boolean

Public Sub SaveLastPosition(AppName As String, ParamKey As String, Object)

    If TypeOf Object Is Form Then
    
        If Not Object.WindowState = vbMinimized Then
            SaveSetting AppName, ParamKey, "Left", Object.Left
            SaveSetting AppName, ParamKey, "Top", Object.Top
            If Object.BorderStyle = 2 Or Object.BorderStyle = 5 Then
                SaveSetting AppName, ParamKey, "Width", Object.Width
                SaveSetting AppName, ParamKey, "Height", Object.Height
            End If
        End If
    
    Else
        SaveSetting AppName, ParamKey, "Left", Object.Left
        SaveSetting AppName, ParamKey, "Top", Object.Top
        If Object.BorderStyle = 2 Or Object.BorderStyle = 5 Then
            SaveSetting AppName, ParamKey, "Width", Object.Width
            SaveSetting AppName, ParamKey, "Height", Object.Height
        End If
    End If
    
End Sub

Public Sub GetLastPosition(AppName As String, ParamKey As String, Object, Optional CenterForm As Boolean = True)
    
    If Not Object.BorderStyle = 2 And Not Object.BorderStyle = 5 Then
        propSaveWindowSize = False
    End If
    
    If Not CenterForm Then
        If Not propSaveWindowSize Then
            Object.Move GetSetting(AppName, ParamKey, "Left", Object.Left), GetSetting(AppName, ParamKey, "Top", Object.Top)
        Else
            Object.Move GetSetting(AppName, ParamKey, "Left", Object.Left), GetSetting(AppName, ParamKey, "Top", Object.Top), GetSetting(AppName, ParamKey, "Width", Object.Width), GetSetting(AppName, ParamKey, "Height", Object.Height)
        End If
    Else
        If Not propSaveWindowSize Then
            Object.Move GetSetting(AppName, ParamKey, "Left", (Screen.Width / 2) - (Object.Width / 2)), GetSetting(AppName, ParamKey, "Top", (Screen.Height / 2) - (Object.Height / 2))
        Else
            Object.Move GetSetting(AppName, ParamKey, "Left", (Screen.Width / 2) - (Object.Width / 2)), GetSetting(AppName, ParamKey, "Top", (Screen.Height / 2) - (Object.Height / 2)), GetSetting(AppName, ParamKey, "Width", Object.Width), GetSetting(AppName, ParamKey, "Height", Object.Height)
        End If
    End If

End Sub

Public Property Let SaveWindowSize(vData As Boolean)
    propSaveWindowSize = vData
End Property

Public Property Get SaveWindowSize() As Boolean
    SaveWindowSize = propSaveWindowSize
End Property

Private Sub Class_Initialize()

    propSaveWindowSize = True
    
End Sub

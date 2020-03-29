VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LaunchPadItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Icon As GDIPImage
Public IconPath As String

'Public Caption As String
Private m_caption As String
Public SpacePositions

Public TargetPath As String
Public TargetArguements As String

Public X As Long
Public Y As Long

Public AlphaIconStore As AlphaIcon

Public GlobalIdentifer As String
Public SearchIdentifier As String

Public Function Clone() As LaunchPadItem
    Set Clone = CloneViItem(Me)
End Function

Public Property Get Caption() As String
    Caption = m_caption
End Property

Public Property Let Caption(newCaption As String)
    m_caption = newCaption
    
    SpacePositions = GetSpacePositions(m_caption)
    SearchIdentifier = UCase(newCaption)
End Property


Private Sub Class_Initialize()
    GlobalIdentifer = GetNextLaunchPadID()
End Sub

Public Function ReleaseIcon()
    Set Icon = Nothing
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeaderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public text As String
Public DisplayedText As String

Public ItemIndex As Long

Private m_Origin As RECTF

Public Function SetOrigin(Height As Single, Left As Single, Top As Single, Width As Single)
    With m_Origin
        .Height = Height
        .Left = Left
        .Top = Top
        .Width = Width
    End With
End Function

Public Function GetOrigin() As RECTF
    GetOrigin = m_Origin
End Function

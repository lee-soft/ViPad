VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPStringFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_format As Long

Private Sub Class_Initialize()
    GdipCreateStringFormat 0, 0, m_format
End Sub

Public Property Let Bold(newBold As Boolean)
    
End Property

Private Sub Class_Terminate()
    If m_format <> 0 Then
        GdipDeleteStringFormat m_format
    End If
End Sub

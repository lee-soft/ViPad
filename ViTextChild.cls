VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FontFace As String
Public Size As Single
Public Y As Long
Public X As Long
Public StringFormat As Long
Public Height As Long
Public Width As Long

Private m_caption As String
Private m_notifyChange As Boolean

Public Property Let Alignment(ByVal newAlignment As GdiPlus.StringAlignment)
    GdipSetStringFormatAlign StringFormat, CLng(newAlignment)
End Property

Public Property Get CalculateWidth() As Long
Dim tempFont As New GDIPFont
Dim tempFontF As New GDIPFontFamily: tempFontF.Constructor FontFace

    tempFont.Constructor tempFontF, Size, FontStyleRegular

    CalculateWidth = ViComponentHelper.MeasureString(m_caption, tempFont)
End Property

Friend Function Changed() As Boolean
    If m_notifyChange Then
        Changed = True
        m_notifyChange = False
    End If
End Function

Public Property Let Caption(newCaption As String)
    If m_caption = newCaption Then
        Exit Property
    End If
    
    m_caption = newCaption
    m_notifyChange = True
End Property

Public Property Get Caption() As String
    Caption = m_caption
End Property

Private Sub Class_Initialize()
    GdipCreateStringFormat 0, 0, StringFormat
    GdipSetStringFormatAlign StringFormat, StringAlignmentNear
End Sub

Private Sub Class_Terminate()
    If StringFormat <> 0 Then
        GdipDeleteStringFormat StringFormat
    End If
End Sub

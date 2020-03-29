VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_unTickedCheckImage As GDIPImage
Private m_tickedCheckBox As GDIPImage

Private m_checked As Boolean
Private m_redrawRequested As Boolean

Public Y As Long
Public X As Long

Public Event onChange()

Public Property Get Checked() As Boolean
    Checked = m_checked
End Property

Public Property Let Checked(newValue As Boolean)
    If m_checked <> newValue Then
        m_checked = newValue

        m_redrawRequested = True
        
        RaiseEvent onChange
    End If
End Property

Public Function RedrawRequest() As Boolean
    RedrawRequest = m_redrawRequested
    m_redrawRequested = False
End Function

Public Function onMouseMove(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function onMouseUp(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function onMouseOut()
End Function

Public Function onMouseDown(MouseButton As Long, X As Single, Y As Single)
    m_checked = Not m_checked
    m_redrawRequested = True
    
    RaiseEvent onChange
End Function

Public Property Get Dimensions_Serialized() As String
    Dimensions_Serialized = Serialize_RectL(Me.GenerateObjectRect)
End Property

Public Function GenerateObjectRect() As RECTL
    
    With GenerateObjectRect
        .Top = Y
        .Left = X
        .Height = m_unTickedCheckImage.Height
        .Width = m_unTickedCheckImage.Width
    End With
    
End Function

Public Function Draw(ByRef theGraphics As GDIPGraphics)

    If m_checked Then
        theGraphics.DrawImage m_tickedCheckBox, CSng(X), CSng(Y), m_tickedCheckBox.Width, m_tickedCheckBox.Height
    Else
        theGraphics.DrawImage m_unTickedCheckImage, CSng(X), CSng(Y), m_unTickedCheckImage.Width, m_unTickedCheckImage.Height
    End If
    
End Function

Private Sub Class_Initialize()

    Set m_unTickedCheckImage = New GDIPImage
    Set m_tickedCheckBox = New GDIPImage
    
    m_unTickedCheckImage.FromBinary LoadResData("EMPTY_CHECKBOX", "IMAGE")
    m_tickedCheckBox.FromBinary LoadResData("TICKED_CHECKBOX", "IMAGE")
    

End Sub

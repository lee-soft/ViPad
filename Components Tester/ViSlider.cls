VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SLOT_TOP_MARGIN As Long = 7

Private m_thisImage As GDIPImage
Private m_theKnob As GDIPImage

Private m_knobX As Long
Private m_sliding As Boolean
Private m_xBuffer As Long

Private m_Max As Long
Private m_Min As Long
Private m_difference As Long


Private m_Y As Long
Private m_X As Long
Private m_width As Long
Private m_redrawRequested As Boolean

Public Event onChange(newValue As Long)

Public Function RedrawRequest() As Boolean
    RedrawRequest = m_redrawRequested
    m_redrawRequested = False
End Function

Public Property Get Dimensions_Serialized() As String
    Dimensions_Serialized = Serialize_RectL(Me.GenerateObjectRect)
End Property

Public Function GenerateObjectRect() As RECTL

Dim returnRect As RECTL
    
    With returnRect
        .Top = m_Y
        .Left = m_X
        .Height = m_theKnob.Height
        .Width = m_width
    End With
    
    GenerateObjectRect = returnRect
    
End Function

Public Property Let Width(newWidth As Long)
    m_width = newWidth
End Property

Public Function onMouseOut()
End Function

Public Function onMouseDown(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function onMouseUp(MouseButton As Long, X As Single, Y As Single)
    m_sliding = False
End Function

Public Function onMouseMove(MouseButton As Long, X As Single, Y As Single)

Dim mousePoint As win.POINTL
    mousePoint.X = CLng(X)
    mousePoint.Y = CLng(Y)
    
    If MouseButton = vbLeftButton Then
        If m_sliding Then
        
            If mousePoint.X + m_xBuffer > (m_width - m_theKnob.Width + 2) Then
                m_knobX = (m_width - m_theKnob.Width + 1)
            ElseIf mousePoint.X + m_xBuffer < -1 Then
                m_knobX = 0
            Else
                m_knobX = CLng(X) + m_xBuffer
            End If
            
        ElseIf RectHelper.PointInsideOfRect(mousePoint, CreateRect(m_theKnob.Height, m_knobX + m_theKnob.Width, m_knobX, 0)) Then
            m_sliding = True
            m_xBuffer = m_knobX - X
            
            If mousePoint.X + m_xBuffer < (m_width - m_theKnob.Width + 2) And _
                mousePoint.X + m_xBuffer > -1 Then m_knobX = CLng(X) + m_xBuffer
        End If
        
        m_redrawRequested = True
        RaiseEvent onChange(Me.Value)
    End If
End Function

Public Property Let Y(newY As Long)
    m_Y = newY
End Property

Public Property Let X(newX As Long)
    m_X = newX
End Property

Public Property Let Max(newMax As Long)
    m_Max = newMax
    m_difference = m_Max - m_Min
End Property

Public Property Let Min(newMin As Long)
    m_Min = newMin
    m_difference = m_Max - m_Min
End Property

Public Property Get Value() As Long
    Value = ((m_difference / (m_width - 27)) * m_knobX) + m_Min
End Property

Public Property Let Value(ByVal newValue As Long)
    newValue = newValue - m_Min
    m_knobX = newValue / (m_difference / (m_width - 27))
End Property

Public Function Draw(ByRef theGraphics As GDIPGraphics)

    theGraphics.DrawImageRect m_thisImage, m_X + 4, m_Y + SLOT_TOP_MARGIN, 4, m_thisImage.Height, 0, 0
    theGraphics.DrawImageRect m_thisImage, m_X + (m_width - 10), m_Y + SLOT_TOP_MARGIN, 4, m_thisImage.Height, 15, 0
    
    DrawImageStretchRect theGraphics, m_thisImage, CreateRectL(m_thisImage.Height, m_width - 18, m_X + 8, m_Y + SLOT_TOP_MARGIN), _
                                        CreateRectL(m_thisImage.Height, 4, 4, 0)
    
    theGraphics.DrawImage m_theKnob, m_X + m_knobX, m_Y + 0, m_theKnob.Width, m_theKnob.Height
End Function

Private Function DrawImageStretchRect(ByRef theGraphics As GDIPGraphics, ByRef Image As GDIPImage, ByRef destRect As RECTL, ByRef sourceRect As RECTL)
    theGraphics.DrawImageStretchAttrF Image, _
        RECTLtoF(destRect), _
        sourceRect.Left, sourceRect.Top, sourceRect.Width, sourceRect.Height, UnitPixel, 0, 0, 0
End Function


Private Sub Class_Initialize()

    Set m_theKnob = New GDIPImage
    Set m_thisImage = New GDIPImage
    
    m_thisImage.FromBinary LoadResData("SLIDER", "IMAGE")
    m_theKnob.FromBinary LoadResData("SLIDER_KNOB", "IMAGE")

End Sub

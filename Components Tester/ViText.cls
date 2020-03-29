VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViText1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_caption As String
Private m_font As GDIPFont
Private m_brush As GDIPBrush

Private m_fontF As GDIPFontFamily
Private m_path As GDIPGraphicPath
Private m_pgb As GDIPPathGradientBrush

Private m_ellipse As RECTF

Private m_width As Long
Private m_Y As Long
Private m_X As Long

Private m_textSize As Single

Private m_redrawRequested As Boolean

Public Property Let Size(newSize As Long)
    m_textSize = CSng(newSize)
    
    m_font.Constructor "Tahoma", m_textSize, FontStyleRegular
    ReconstructPaths
End Property

Public Function RedrawRequest() As Boolean
    RedrawRequest = m_redrawRequested
    m_redrawRequested = False
End Function

Private Function ReconstructPaths()
Dim A As Long
    Set m_path = New GDIPGraphicPath

    GdipCreateStringFormat 0, 0, A
    GdipSetStringFormatAlign A, StringAlignmentNear
    GdipSetStringFormatLineAlign A, StringAlignmentNear
    
    m_width = MeasureString(m_caption, m_font)
    m_path.AddString m_caption, m_fontF, fontStyle.FontStyleRegular, m_textSize, CreateRectF(m_X, m_Y, CLng(m_textSize), m_X + m_width), A
    
    GdipDeleteStringFormat A
        
Dim centerColour As Colour
Dim suroundColour() As Colour
Dim tempPath As GDIPGraphicPath
    
    Set centerColour = New Colour
    centerColour.SetColourByHex "ffffff"
    
    ReDim suroundColour(0)
    Set suroundColour(0) = New Colour
    suroundColour(0).SetColourByHex "ffffff"
    suroundColour(0).Alpha = 5

    Set tempPath = New GDIPGraphicPath
    tempPath.AddEllipse CreateRectF(m_X - m_textSize, m_Y + 3, CLng(m_textSize), m_width + (m_textSize * 2))
    
    Set m_pgb = New GDIPPathGradientBrush
    m_pgb.FromPath tempPath
    m_pgb.centerColour = centerColour
    m_pgb.SurroundColours suroundColour
    
    m_redrawRequested = True
End Function

Public Property Let Caption(newCaption As String)
    m_caption = newCaption
    
    ReconstructPaths
End Property

Public Property Let Y(newY As Long)
    m_Y = newY
    ReconstructPaths
End Property

Public Property Let X(newX As Long)
    m_X = newX
    ReconstructPaths
End Property

Public Property Get Dimensions_Serialized() As String
    Dimensions_Serialized = Serialize_RectL(Me.GenerateObjectRect)
End Property

Public Function GenerateObjectRect() As RECTL

Dim returnRect As RECTL
    
    With returnRect
        .Top = m_Y
        .Left = m_X
        .Height = 16
        .Width = m_width
    End With
    
    GenerateObjectRect = returnRect
    
End Function

Public Function onMouseMove(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function onMouseDown(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function onMouseUp(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function Draw(ByRef theGraphics As GDIPGraphics)
    
    theGraphics.FillEllipse m_pgb, CreateRectF(m_X, m_Y, CLng(m_textSize), m_width)
    theGraphics.FillPath m_brush, m_path

    'theGraphics.DrawString m_caption, m_font, m_brushShadow, m_positionShadow
    'theGraphics.DrawString m_caption, m_font, m_brush, m_position

End Function

Private Sub Class_Initialize()

    Set m_brush = New GDIPBrush
    Set m_path = New GDIPGraphicPath
    
    m_brush.Colour.SetColourByHex "000000"
    
    Set m_font = New GDIPFont
    Set m_fontF = New GDIPFontFamily
    
    m_textSize = 20
    
    m_font.Constructor "Tahoma", m_textSize, FontStyleRegular
    m_fontF.Constructor "Tahoma"

    m_caption = "<Empty String>"
    
    m_ellipse.Width = 100
    m_ellipse.Height = 10
    
    Me.X = 0
    Me.Y = 0
End Sub

Private Function GenerateEllipse()
    
End Function

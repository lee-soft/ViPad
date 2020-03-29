VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViTextLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BLUR_SCALE As Single = 5

Private m_width As Long
Private m_Y As Long
Private m_X As Long

Private m_textSize As Single
Private m_redrawRequested As Boolean

Private m_children As Collection

Private m_pth As GDIPGraphicPath
Private m_bitmap As GDIPBitmap
Private m_form As Form

Private Function HasChildrenChanged() As Boolean

Dim thisText As ViText

    For Each thisText In m_children
        If thisText.Changed Then
            'Dont exit, because it resets all other children
            HasChildrenChanged = True
        End If
    Next
End Function

Public Function CreateChild(ByVal szText As String, ByVal X As Long, ByVal Y As Long, Optional szFontFace As String = "Tahoma", Optional fontSize As Long = 9) As ViText

Dim thisChild As New ViText
    m_children.Add thisChild
    
    thisChild.Caption = szText
    thisChild.FontFace = szFontFace
    thisChild.Size = fontSize
    thisChild.X = X
    thisChild.Y = Y
    
    Set CreateChild = thisChild
End Function

Public Property Let Parent(newForm As Form)
    Set m_form = newForm
End Property

Public Function RedrawRequest() As Boolean
    RedrawRequest = m_redrawRequested
    m_redrawRequested = False
End Function

Private Function ReconstructPaths()

Dim thisChild As ViText

    If m_form Is Nothing Then Exit Function

    Set m_bitmap = New GDIPBitmap: m_bitmap.CreateFromSizeFormat m_form.ScaleWidth / BLUR_SCALE, m_form.ScaleHeight / BLUR_SCALE, PixelFormat.Format32bppArgb
    Set m_pth = New GDIPGraphicPath: m_pth.Constructor FillModeWinding
    
    Dim thisFontF As GDIPFontFamily

    For Each thisChild In m_children
        Set thisFontF = New GDIPFontFamily
        thisFontF.Constructor (thisChild.FontFace)
        
        m_pth.AddString thisChild.Caption, CreateFontFamily(thisChild.FontFace), FontStyle.FontStyleRegular, thisChild.Size, CreateRectF(thisChild.X, thisChild.Y, thisChild.Height, thisChild.Width), thisChild.StringFormat
    Next
    
    
    Dim G As New GDIPGraphics: G.FromImage m_bitmap.Image
    Dim mx As New GDIPMatrix: mx.Constructor2 1# / BLUR_SCALE, 0, 0, 1# / BLUR_SCALE, 1# / BLUR_SCALE, 1# / BLUR_SCALE
                                             '1.0f/5,0,0,1.0f/5,-(1.0f/5),-(1.0f/5)

    G.SmoothingMode = SmoothingModeHighQuality
    G.Transform mx

    Dim p As New GDIPPen: p.Constructor CreateWebColour("FFFFF2"), 1, 128
    G.DrawPath p, m_pth
    G.FillPath Custom_Brush(CreateWebColour("FFFFF2")), m_pth
    G.Dispose
    
End Function

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

    If m_form Is Nothing Then Exit Function
    If m_bitmap Is Nothing Then ReconstructPaths

    If HasChildrenChanged Then
        ReconstructPaths
    End If

    theGraphics.DrawImageStretchAttrL m_bitmap.Image, _
        CreateRectL(m_form.ScaleHeight, m_form.ScaleWidth, 0, 0), _
        0, 0, m_bitmap.Image.Width, m_bitmap.Image.Height, UnitPixel, 0, 0, 0
    
    'theGraphics.DrawImageStretchAttrL m_bitmap.Image, _
        CreateRectL(m_form.ScaleHeight, m_form.ScaleWidth, 0, 0), _
        0, 0, m_bitmap.Image.Width, m_bitmap.Image.Height, UnitPixel, 0, 0, 0
        
    'theGraphics.DrawImageStretchAttrL m_bitmap.Image, _
        CreateRectL(m_form.ScaleHeight, m_form.ScaleWidth, 0, 0), _
        0, 0, m_bitmap.Image.Width, m_bitmap.Image.Height, UnitPixel, 0, 0, 0
    
    theGraphics.FillPath Brushes_Black, m_pth

End Function

Private Sub Class_Initialize()
    Set m_children = New Collection
End Sub

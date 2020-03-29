VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_currentBitmap As GDIPBitmap
Private m_font As GDIPFont
Private m_fontF As GDIPFontFamily
Private m_thisImage As GDIPImage
Private m_thisButton2 As GDIPImage

Private m_brush As GDIPBrush
Private m_path As GDIPGraphicPath

Private m_rectPosition As RECTF
Private m_clicked As Boolean
Private m_width As Long

Private m_caption As String
Private m_command As String

Private m_redrawRequested As Boolean

Private m_Y As Long
Private m_X As Long

Public Event onClicked()

Public Property Get Dimensions_Serialized() As String
    Dimensions_Serialized = Serialize_RectL(Me.GenerateObjectRect)
End Property

Public Function GenerateObjectRect() As RECTL

Dim returnRect As RECTL
    
    With returnRect
        .Top = m_Y
        .Left = m_X
        .Height = m_thisImage.Height
        .Width = m_width
    End With
    
    GenerateObjectRect = returnRect
    
End Function

Public Function RedrawRequest() As Boolean
    RedrawRequest = m_redrawRequested
    m_redrawRequested = False
End Function

Public Function onMouseOut()
    If m_clicked Then
        m_clicked = False
        
        m_redrawRequested = True
    End If
End Function

Public Function onMouseMove(MouseButton As Long, X As Single, Y As Single)
End Function

Public Function onMouseDown(MouseButton As Long, X As Single, Y As Single)
    m_clicked = True
    m_redrawRequested = True
End Function

Public Function onMouseUp(MouseButton As Long, X As Single, Y As Single)
    m_clicked = False
    m_redrawRequested = True
    
    RaiseEvent onClicked
End Function

Public Property Let Width(newWidth As Long)
    m_width = newWidth
    
    m_rectPosition.Width = m_width - 5
End Property

Public Property Let Y(newY As Long)
    m_Y = newY
    
    m_rectPosition.Top = m_Y + 3
    
    MakePath
End Property

Public Property Let X(newX As Long)
    m_X = newX
    
    m_rectPosition.Left = m_X + 3
    
    MakePath
End Property

Private Function MakePath()

    If m_caption = vbNullString Then
        Exit Function
    End If

Dim theFontStyle As Long
Dim A As Long

    GdipCreateStringFormat 0, 0, A
    GdipSetStringFormatAlign A, StringAlignmentCenter
    GdipSetStringFormatLineAlign A, StringAlignmentCenter

    Set m_path = New GDIPGraphicPath
    theFontStyle = FontStyle.FontStyleRegular

    m_path.AddString m_caption, m_fontF, theFontStyle, 15, m_rectPosition, A
    GdipDeleteStringFormat A
    
End Function

Private Sub Class_Initialize()

    Set m_currentBitmap = New GDIPBitmap
    Set m_path = New GDIPGraphicPath
    Set m_fontF = New GDIPFontFamily
    Set m_thisImage = New GDIPImage
    Set m_thisButton2 = New GDIPImage
    
    Set m_font = New GDIPFont
    Set m_brush = New GDIPBrush

    m_brush.Colour.Value = vbBlack
    m_fontF.Constructor "Tahoma"
    
    Me.Width = 200

    m_thisImage.FromBinary LoadResData("BUTTON", "IMAGE")
    m_thisButton2.FromBinary LoadResData("BUTTON2", "IMAGE")

    m_rectPosition.Height = m_thisImage.Height - (13 / 2)
End Sub

Public Property Let Caption(ByVal newCaption As String)
    m_caption = newCaption
    
    MakePath
End Property

Public Function Draw(ByRef theGraphics As GDIPGraphics)

Dim labelPoint As POINTF
Dim image2Draw As GDIPImage
    
    If m_clicked Then
        Set image2Draw = m_thisImage
    Else
        Set image2Draw = m_thisButton2
    End If
    
    theGraphics.DrawImageRect image2Draw, m_X, m_Y, 5, m_thisImage.Height, 0, 0
    theGraphics.DrawImageRect image2Draw, m_X + (m_width - 5), m_Y, 5, m_thisImage.Height, 8, 0
    DrawImageStretchRect theGraphics, image2Draw, CreateRectL(m_thisImage.Height, m_width - 10, m_X + 5, m_Y), _
                                        CreateRectL(m_thisImage.Height, 4, 5, 0)
    
    theGraphics.FillPath m_brush, m_path
End Function

Private Function DrawImageStretchRect(ByRef theGraphics As GDIPGraphics, ByRef Image As GDIPImage, ByRef destRect As RECTL, ByRef sourceRect As RECTL)
    theGraphics.DrawImageStretchAttrF Image, _
        RECTLtoF(destRect), _
        sourceRect.Left, sourceRect.Top, sourceRect.Width, sourceRect.Height, UnitPixel, 0, 0, 0
End Function


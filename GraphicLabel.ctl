VERSION 5.00
Begin VB.UserControl GraphicLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "GraphicLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_graphics As GDIPGraphics
Private m_graphicsImage As GDIPGraphics
Private m_currentBitmap As GDIPBitmap
Private m_font As GDIPFont
Private m_fontF As GDIPFontFamily

Private m_brush As GDIPBrush
Private m_brushShadow As GDIPBrush

Private m_path As GDIPGraphicPath
Private m_pathShadow As GDIPGraphicPath
Private m_rectPosition As RECTF

Private m_caption As String
Private m_autoSize As Boolean
Private m_multiLine As Boolean

Public Property Let MultiLine(newValue As Boolean)
    m_multiLine = newValue
    AutoSizeControl
End Property

Private Function MakePath()

Dim theFontStyle As Long
Dim rectShadow As RECTF

    Set m_path = New GDIPGraphicPath
    Set m_pathShadow = New GDIPGraphicPath
    
    If UserControl.FontBold Then
        theFontStyle = FontStyle.FontStyleBold
    ElseIf UserControl.FontItalic Then
        theFontStyle = FontStyle.FontStyleItalic
    Else
        theFontStyle = FontStyle.FontStyleRegular
    End If
    
    Dim m_fontF As New GDIPFontFamily: m_fontF.Constructor (UserControl.Font.Name)

    m_font.Constructor m_fontF, UserControl.Font.Size, theFontStyle
    m_path.AddString m_caption, m_fontF, theFontStyle, UserControl.Font.Size, m_rectPosition, 0
    
    rectShadow.Left = m_rectPosition.Left + 1
    rectShadow.Top = m_rectPosition.Top + 1
    rectShadow.Height = m_rectPosition.Height
    rectShadow.Width = m_rectPosition.Width
    
    m_pathShadow.AddString m_caption, m_fontF, theFontStyle, UserControl.Font.Size, rectShadow, 0
End Function

Private Function AutoSizeControl()

Dim thisLayout As RECTF
Dim theDimensions As RECTF

    If m_multiLine Then
        thisLayout.Width = UserControl.ScaleWidth
    End If

    theDimensions = m_graphicsImage.MeasureStringEx(m_caption, m_font, 0, thisLayout)
    
    UserControl.Height = Ceiling(CDbl(theDimensions.Height)) * Screen.TwipsPerPixelY
    UserControl.Width = Ceiling(CDbl(theDimensions.Width) + 1) * Screen.TwipsPerPixelX
    
    'Debug.Print Ceiling(theDimensions.Width * Screen.TwipsPerPixelX)
End Function

Public Property Let AutoSize(newValue As Boolean)
    m_autoSize = newValue

    If m_autoSize Then
        AutoSizeControl
    End If
    
    PropertyChanged "Autosize"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_autoSize
End Property

Public Property Let Caption(newCaption As String)
    m_caption = newCaption
    
    If m_autoSize Then
        AutoSizeControl
    End If
    
    MakePath
    DrawSomething
End Property

Public Property Get Caption() As String
    Caption = m_caption
End Property

Private Function DrawSomething()
    If InIDE Then Exit Function

Dim labelPoint As POINTF

    If Ambient.UserMode Then
        If UserControl.BackColor = vbBlack Then
            m_graphicsImage.Clear
        Else
            m_graphicsImage.Clear UserControl.BackColor
        End If
    Else
        m_graphicsImage.Clear vbWhite
    End If
    
    If UserControl.BackColor <> vbWhite Then m_graphicsImage.FillPath m_brushShadow, m_pathShadow
    m_graphicsImage.FillPath m_brush, m_path
    
    RefreshHDC
End Function

Function DrawImageStretchRect(ByRef Image As GDIPImage, ByRef destRect As RECTL, ByRef sourceRect As RECTL)
    m_graphicsImage.DrawImageStretchAttrF Image, _
        RECTLtoF(destRect), _
        sourceRect.Left, sourceRect.Top, sourceRect.Width, sourceRect.Height, UnitPixel, 0, 0, 0
End Function

Private Function InitializeGraphics()
    Set m_graphics = New GDIPGraphics
    m_graphics.FromHDC UserControl.hdc
    
    m_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.InterpolationMode = InterpolationModeHighQualityBicubic
End Function

Private Sub UserControl_Initialize()

    Dim colorMaker As New Colour
    colorMaker.SetColourByHex "#363535"

    InitializeGDIIfNotInitialized

    Set m_graphics = New GDIPGraphics
    Set m_graphicsImage = New GDIPGraphics
    Set m_currentBitmap = New GDIPBitmap
    Set m_path = New GDIPGraphicPath
    Set m_fontF = New GDIPFontFamily
    
    Set m_font = New GDIPFont
    Set m_brush = New GDIPBrush
    Set m_brushShadow = New GDIPBrush

    m_fontF.Constructor UserControl.Font.Name
    m_font.Constructor m_fontF, UserControl.Font.Size, FontStyleRegular
    
    m_brush.Colour.Value = UserControl.ForeColor
    m_brushShadow.Colour.Value = colorMaker.Value

    
    InitializeGraphics
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InIDE Then Exit Sub

    DrawSomething
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "{No Caption}")
    
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    
    m_brush.Colour.Value = UserControl.ForeColor
    
    m_fontF.Constructor UserControl.Font.Name
    m_font.Constructor m_fontF, UserControl.Font.Size, FontStyleRegular
    
    AutoSize = PropBag.ReadProperty("Autosize", False)
    
    MakePath
    DrawSomething
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_caption, "{No Caption}")
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Autosize", m_autoSize, False)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

Private Sub UserControl_Resize()
    If InIDE Then Exit Sub
    
    If m_autoSize Then
        AutoSizeControl
    End If
    
    m_currentBitmap.CreateFromSizeFormat UserControl.ScaleWidth, UserControl.ScaleHeight, PixelFormat.Format32bppArgb
    m_graphicsImage.FromImage m_currentBitmap.Image
    m_graphicsImage.SmoothingMode = SmoothingModeHighQuality
    m_graphicsImage.InterpolationMode = InterpolationModeHighQualityBicubic
    
    m_rectPosition.Width = UserControl.ScaleWidth
    m_rectPosition.Height = UserControl.ScaleHeight

    InitializeGraphics
    MakePath
    DrawSomething
End Sub

Private Sub RefreshHDC()
    'On Error GoTo Handler
    If m_graphics Is Nothing Then Exit Sub
    
    m_graphics.Clear
    m_graphics.DrawImage m_currentBitmap.Image, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    'm_graphics.FillPath m_brush, m_path

Handler:
    UserControl.Refresh
End Sub

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    
    Debug.Print "Set Font:: " & Font.Size
    
    m_fontF.Constructor UserControl.Font.Name
    m_font.Constructor m_fontF, UserControl.Font.Size, FontStyleRegular
    
    If m_autoSize Then
        AutoSizeControl
    End If
    
    MakePath
    DrawSomething
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    
    m_brush.Colour.Value = UserControl.ForeColor
    DrawSomething
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property


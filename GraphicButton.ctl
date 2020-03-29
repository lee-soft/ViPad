VERSION 5.00
Begin VB.UserControl GraphicButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
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
Attribute VB_Name = "GraphicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event onClick()

Private m_graphics As GDIPGraphics
Private m_graphicsImage As GDIPGraphics
Private m_currentBitmap As GDIPBitmap
Private m_font As GDIPFont
Private m_fontF As GDIPFontFamily
Private m_thisImage As GDIPImage
Private m_thisButton2 As GDIPImage

Private m_brush As GDIPBrush
Private m_path As GDIPGraphicPath
Private m_rectPosition As RECTF
Private m_clicked As Boolean

Private m_caption As String
Private m_command As String

Public Property Let Command(ByVal newCommand As String)
    m_command = newCommand
End Property

Public Property Get Command() As String
    Command = m_command
End Property

Private Function MakePath()

Dim theFontStyle As Long
Dim A As Long

    GdipCreateStringFormat 0, 0, A
    GdipSetStringFormatAlign A, StringAlignmentCenter
    GdipSetStringFormatLineAlign A, StringAlignmentCenter

    Set m_path = New GDIPGraphicPath
    
    If UserControl.FontBold Then
        theFontStyle = FontStyle.FontStyleBold
    ElseIf UserControl.FontItalic Then
        theFontStyle = FontStyle.FontStyleItalic
    Else
        theFontStyle = FontStyle.FontStyleRegular
    End If
    
    m_path.AddString m_caption, m_fontF, theFontStyle, UserControl.Font.Size, m_rectPosition, A
    
    GdipDeleteStringFormat A
    
End Function

Public Property Let Caption(newCaption As String)
    m_caption = newCaption
    
    MakePath
    DrawSomething
End Property

Public Property Get Caption() As String
    Caption = m_caption
End Property

Private Function DrawSomething()

Dim labelPoint As POINTF
Dim image2Draw As GDIPImage
    
    If m_clicked Then
        Set image2Draw = m_thisImage
    Else
        Set image2Draw = m_thisButton2
    End If

    If Ambient.UserMode Then
        If UserControl.BackColor = vbBlack Then
            m_graphicsImage.Clear
        Else
            m_graphicsImage.Clear UserControl.BackColor
        End If
    Else
        m_graphicsImage.Clear vbWhite
    End If
    
    m_graphicsImage.DrawImageRect image2Draw, 0, 0, 5, m_thisImage.Height, 0, 0
    m_graphicsImage.DrawImageRect image2Draw, UserControl.ScaleWidth - 5, 0, 5, m_thisImage.Height, 8, 0
    DrawImageStretchRect image2Draw, CreateRectL(m_thisImage.Height, UserControl.ScaleWidth - 10, 5, 0), _
                                        CreateRectL(m_thisImage.Height, 4, 5, 0)
    
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

Private Sub UserControl_Click()
    RaiseEvent onClick
End Sub

Private Sub UserControl_Initialize()

    InitializeGDIIfNotInitialized

    Set m_graphics = New GDIPGraphics
    Set m_graphicsImage = New GDIPGraphics
    Set m_currentBitmap = New GDIPBitmap
    Set m_path = New GDIPGraphicPath
    Set m_fontF = New GDIPFontFamily
    Set m_thisImage = New GDIPImage
    Set m_thisButton2 = New GDIPImage
    
    Set m_font = New GDIPFont
    Set m_brush = New GDIPBrush
    
    m_thisImage.FromBinary LoadResData("BUTTON", "IMAGE")
    m_thisButton2.FromBinary LoadResData("BUTTON2", "IMAGE")
    'm_thisImage.FromResource "BUTTON", "IMAGE"
    'm_thisButton2.FromResource "BUTTON2", "IMAGE"
    
    m_fontF.Constructor UserControl.Font.Name
    m_brush.Colour.Value = UserControl.ForeColor
    
    InitializeGraphics
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InIDE Then Exit Sub
    
    m_clicked = True
    DrawSomething
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InIDE Then Exit Sub
    
    m_clicked = False
    DrawSomething
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If InIDE Then Exit Sub
    
    Caption = PropBag.ReadProperty("Caption", "{No Caption}")
    
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    
    m_brush.Colour.Value = UserControl.ForeColor
    m_fontF.Constructor UserControl.Font.Name
    
    MakePath
    DrawSomething
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H0&)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", m_caption, "{No Caption}"
    
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H0&)
End Sub


Private Sub UserControl_Resize()
    If InIDE Then Exit Sub
    
    m_currentBitmap.CreateFromSizeFormat UserControl.ScaleWidth, UserControl.ScaleHeight, PixelFormat.Format32bppArgb
    m_graphicsImage.FromImage m_currentBitmap.Image
    m_graphicsImage.SmoothingMode = SmoothingModeHighQuality
    m_graphicsImage.InterpolationMode = InterpolationModeHighQualityBicubic
    
    m_rectPosition.Top = 0
    m_rectPosition.Left = 2
    
    m_rectPosition.Width = UserControl.ScaleWidth - 5
    m_rectPosition.Height = UserControl.ScaleHeight
    
    UserControl.Height = m_thisImage.Height * Screen.TwipsPerPixelY

    InitializeGraphics
    MakePath
    DrawSomething
End Sub

Private Sub RefreshHDC()
    'On Error GoTo Handler
    If m_graphics Is Nothing Then Exit Sub
    
    If UserControl.BackColor = vbBlack Then
        m_graphics.Clear
    Else
        m_graphics.Clear UserControl.BackColor
    End If
    
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
    
    UserControl_Resize
End Property


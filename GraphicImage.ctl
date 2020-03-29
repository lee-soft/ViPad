VERSION 5.00
Begin VB.UserControl GraphicImage 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "GraphicImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_graphics As GDIPGraphics
Private m_graphicsImage As GDIPGraphics
Private m_currentBitmap As GDIPBitmap

Private m_thisImage As GDIPImage

Public Property Let Image(ByRef newImage As GDIPImage)
    Set m_thisImage = newImage
    
    UserControl.Width = m_thisImage.Width * Screen.TwipsPerPixelX
    UserControl.Height = m_thisImage.Height * Screen.TwipsPerPixelY
    
    DrawSomething
End Property

Private Function DrawSomething()
    m_graphicsImage.Clear
    m_graphicsImage.DrawImage m_thisImage, 0, 0, m_thisImage.Width, m_thisImage.Height

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
    InitializeGDIIfNotInitialized

    Set m_graphics = New GDIPGraphics
    Set m_graphicsImage = New GDIPGraphics
    Set m_currentBitmap = New GDIPBitmap
    
    Set m_thisImage = New GDIPImage
    InitializeGraphics
End Sub

Private Sub UserControl_Resize()
    If InIDE Then Exit Sub

    m_currentBitmap.CreateFromSizeFormat UserControl.ScaleWidth, UserControl.ScaleHeight, PixelFormat.Format32bppArgb
    m_graphicsImage.FromImage m_currentBitmap.Image

    InitializeGraphics
    DrawSomething
    
    'UserControl.width = m_tickedCheckBox.width * Screen.TwipsPerPixelX
    'UserControl.height = m_tickedCheckBox.height * Screen.TwipsPerPixelY
End Sub

Private Sub RefreshHDC()
    On Error GoTo Handler
    If m_graphics Is Nothing Then Exit Sub
    
    m_graphics.Clear
    m_graphics.DrawImage m_currentBitmap.Image, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
Handler:
    UserControl.Refresh
End Sub





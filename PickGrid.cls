VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PickGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_whiteText As GDIPGraphicPath
Private m_blackText As GDIPGraphicPath
Private m_graphicsImage As GDIPGraphics

Private m_currentBitmap As GDIPBitmap

Public Items As Collection

Public Function ClearTextGraphicPaths()
    Set m_blackText = New GDIPGraphicPath
    Set m_whiteText = New GDIPGraphicPath
End Function

Public Property Get GraphicsImage() As GDIPGraphics
    Set GraphicsImage = m_graphicsImage
End Property

Public Property Get WhiteTextGP() As GDIPGraphicPath
    Set WhiteTextGP = m_whiteText
End Property

Public Property Get BlackTextGP() As GDIPGraphicPath
    Set BlackTextGP = m_blackText
End Property

Public Property Get Bitmap() As GDIPBitmap
    Set Bitmap = m_currentBitmap
End Property

Private Sub Class_Initialize()
    Set m_currentBitmap = New GDIPBitmap

    ClearTextGraphicPaths
End Sub

Public Property Let WindowSize(newSize As SIZEL)
    Set m_graphicsImage = New GDIPGraphics
    
    m_currentBitmap.CreateFromSizeFormat newSize.Width, newSize.Height, PixelFormat.Format32bppArgb
    m_graphicsImage.FromImage m_currentBitmap.Image

    m_graphicsImage.SmoothingMode = SmoothingModeHighQuality
    m_graphicsImage.InterpolationMode = InterpolationModeHighQualityBicubic
End Property

Private Sub Class_Terminate()
    m_currentBitmap.Dispose
    m_whiteText.Dispose
    m_blackText.Dispose
    m_currentBitmap.Dispose
    
End Sub


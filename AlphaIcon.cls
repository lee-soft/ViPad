VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_icon As GDIPBitmap
Private m_IconXP As GDIPBitmap

Private m_IsAlphaBitmap As Boolean

Public Property Get Image() As GDIPImage
    Set Image = m_IconXP.Image
End Property

Public Function CreateFromHICON(ByVal icoHandle As Long)

    

Dim ii As ICONINFO

Dim bmData As BitmapData
Dim bmBounds As RECTL

Dim x As Long
Dim y As Long

    Set m_icon = New GDIPBitmap
    Set m_IconXP = New GDIPBitmap

    GetIconInfo icoHandle, ii
    m_icon.CreateFromHBITMAP ii.hbmColor, 0
    
    'm_icon.Image.Save "C:\test2.png", GetPngCodecCLSID

    DeleteObject ii.hbmColor
    DeleteObject ii.hbmMask
        
    If m_icon.Image.ImgPixelFormat = PixelFormat.Format32bppArgb Then
    End If
        

        'If BITMAP.GetPixelFormatSize(m_Icon.PixelFormat) < 32 Then
            'Return ico.ToBitmap
        'End If

    bmBounds.Width = m_icon.Image.Width
    bmBounds.Height = m_icon.Image.Height

    bmData = m_icon.LockBits(bmBounds, _
                ImageLockModeRead, _
                m_icon.Image.ImgPixelFormat)

    m_IconXP.CreateFromSizeFormatData bmData.Height, _
                bmData.Width, _
                bmData.stride, _
                Format32bppArgb, _
                bmData.Scan0Ptr


        
    m_IsAlphaBitmap = False

    For y = 0 To bmData.Height - 1
        For x = 0 To bmData.Width - 1
            
            Dim PixelColor As argb
            Long2ARGB m_IconXP.GetPixel(x, y), PixelColor
                'PixelColor = Color.FromArgb(Marshal.ReadInt32(bmData.Scan0, (bmData.Stride * y) + (4 * x)))
            If PixelColor.A > 0 And PixelColor.A < 255 Then
                m_IsAlphaBitmap = True
                Exit For
            End If
        Next
        If m_IsAlphaBitmap Then Exit For
    Next

    m_icon.UnlockBits bmData
        
    If Not m_IsAlphaBitmap Then
        m_icon.Dispose
        m_IconXP.Dispose
        
        m_IconXP.CreateFromHICON icoHandle
    End If
    
End Function


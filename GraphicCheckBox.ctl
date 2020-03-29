VERSION 5.00
Begin VB.UserControl GraphicCheckBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "GraphicCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_graphics As GDIPGraphics
Private m_graphicsImage As GDIPGraphics
Private m_currentBitmap As GDIPBitmap

Private m_thisImage As GDIPImage
Private m_tickedCheckBox As GDIPImage
Private m_value As Boolean

Public Event onChange()

Public Property Get Value() As Boolean
    Value = m_value
End Property

Public Property Let Value(newValue As Boolean)
    If Not m_value = newValue Then
        m_value = newValue
        DrawSomething
        
        RaiseEvent onChange
    End If
End Property

Private Function DrawSomething()
    If UserControl.BackColor = vbBlack Then
        m_graphicsImage.Clear
    Else
        m_graphicsImage.Clear UserControl.BackColor
    End If
    
    If m_value Then
        m_graphicsImage.DrawImage m_tickedCheckBox, 0, 0, m_tickedCheckBox.Width, m_tickedCheckBox.Height
    Else
        m_graphicsImage.DrawImage m_thisImage, 0, 0, m_thisImage.Width, m_thisImage.Height
    End If
    
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
    Set m_tickedCheckBox = New GDIPImage
    
    m_thisImage.FromBinary LoadResData("EMPTY_CHECKBOX", "IMAGE")
    m_tickedCheckBox.FromBinary LoadResData("TICKED_CHECKBOX", "IMAGE")
    
    InitializeGraphics
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_value = Not m_value
    DrawSomething
    
    RaiseEvent onChange
End Sub

Private Sub UserControl_Resize()
    If InIDE Then Exit Sub
    
    m_currentBitmap.CreateFromSizeFormat UserControl.ScaleWidth, UserControl.ScaleHeight, PixelFormat.Format32bppArgb
    m_graphicsImage.FromImage m_currentBitmap.Image

    InitializeGraphics
    DrawSomething
    
    UserControl.Width = m_tickedCheckBox.Width * Screen.TwipsPerPixelX
    UserControl.Height = m_tickedCheckBox.Height * Screen.TwipsPerPixelY
End Sub

Private Sub RefreshHDC()
    On Error GoTo Handler
    If m_graphics Is Nothing Then Exit Sub
    
    If UserControl.BackColor = vbBlack Then
        m_graphics.Clear
    Else
        m_graphics.Clear UserControl.BackColor
    End If
    
    m_graphics.DrawImage m_currentBitmap.Image, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
Handler:
    UserControl.Refresh
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

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFF&)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFF&)
End Sub


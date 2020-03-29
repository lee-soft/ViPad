VERSION 5.00
Begin VB.Form FloatingImage 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   LinkTopic       =   "Form1"
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   14
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer timFollowCursor 
      Interval        =   100
      Left            =   840
      Top             =   1560
   End
End
Attribute VB_Name = "FloatingImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Private m_graphics As GDIPGraphics
Private m_Image As GDIPImage
Private m_layeredWindow As LayerdWindowHandles

Public FilePath As String

Public Sub SetImage(ByRef newImage As GDIPImage, Optional ByVal imageSize As Long = -1)
    Set m_Image = newImage
    
    If imageSize < 1 Then
        Me.Height = m_Image.Height * Screen.TwipsPerPixelY
        Me.Width = m_Image.Width * Screen.TwipsPerPixelX
    Else
        Me.Height = imageSize * Screen.TwipsPerPixelY
        Me.Width = Me.Height
    End If
    
    Set m_layeredWindow = MakeLayerdWindow(Me, True, True)
    m_graphics.FromHDC m_layeredWindow.theDC
    
    ReDrawImage
    
    Me.Visible = True
End Sub

Private Sub ReDrawImage()
    m_graphics.Clear
    m_graphics.DrawImage m_Image, 0, 0, m_Image.Width, m_Image.Height
    
    UpdateWindow
End Sub

Private Sub Form_Load()
    InitializeGDIIfNotInitialized
    Set m_graphics = New GDIPGraphics
End Sub

Public Function UpdateWindow()
    Debug.Print "GlassWindow::UpdateWindow()"

    If m_layeredWindow Is Nothing Then Exit Function
    m_layeredWindow.Update Me.hWnd, Me.hdc
End Function

Sub MoveToCursor()
    
Dim cursorPosition As win.POINTL

    GetCursorPos cursorPosition
    Me.Move (cursorPosition.X - (Me.ScaleWidth / 2)) * Screen.TwipsPerPixelX, _
            (cursorPosition.Y - (Me.ScaleHeight / 2)) * Screen.TwipsPerPixelY
    
End Sub

Sub TrackCursor()

Dim cursorPosition As win.POINTL

    While (True)
        GetCursorPos cursorPosition
        Me.Move (cursorPosition.X - (Me.ScaleWidth / 2)) * Screen.TwipsPerPixelX, _
                (cursorPosition.Y - (Me.ScaleHeight / 2)) * Screen.TwipsPerPixelY
        
        DoEvents
    Wend

End Sub

Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
    Data.Files.Add FilePath
End Sub

Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    On Error GoTo Handler
    ' Set data format to file.
    Data.SetData , vbCFFiles
    ' Display the move mouse pointer..
    AllowedEffects = vbDropEffectCopy
    
    Data.Files.Add FilePath
Handler:
End Sub

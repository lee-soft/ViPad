VERSION 5.00
Begin VB.Form GlassContainer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13845
   LinkTopic       =   "Form3"
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   923
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   4200
   End
End
Attribute VB_Name = "GlassContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Option Explicit

Private m_layeredWindow As LayerdWindowHandles
Private m_client As Form

Private m_slices As Collection
Private m_graphics As GDIPGraphics

Private m_registeredWidth As Long
Private m_registeredHeight As Long

Dim backgroundImage As GDIPImage

Private m_POS As Long
Private m_hasFocus As Boolean
Private m_clientSize As SIZEL
Private m_clientSlice As Slice
Private m_invalid As Boolean
Private m_oldParenthWnd As Long
Private m_trackingMouse As Boolean

Implements IHookSink


Public Event onResized()
Public Event onMouseMove(Button As Integer, X As Long, Y As Long)
Public Event onMouseUp(Button As Integer, X As Long, Y As Long)
Public Event onMouseDown(Button As Integer, X As Long, Y As Long)
Public Event QueryOnItem(X As Long, Y As Long, ByRef onItem As Boolean)
Public Event onDblClick()

Public Event onKeyPress(KeyAscii As Long)

Public Property Let DragDrop(bNewAccept As Boolean)
    If bNewAccept = True Then
        DragAcceptFiles Me.hWnd, APITRUE
    End If
End Property

Public Function ReleaseClient()
    m_invalid = True
    
    UnhookWindow Me.hWnd
    
    SetWindowLong m_client.hWnd, GWL_HWNDPARENT, m_oldParenthWnd
    Set m_client = Nothing
End Function

Public Function IsCursorOnBoarder() As Boolean

Dim g_paCursorPos As win.POINTL
Dim m_MyPosition As win.RECT

    GetCursorPos g_paCursorPos
    GetWindowRect Me.hWnd, m_MyPosition
    
    If (g_paCursorPos.X < m_MyPosition.Right) And _
        (g_paCursorPos.X > m_MyPosition.Left + Me.ScaleWidth - 22) Then
        
        IsCursorOnBoarder = True
        Exit Function
    End If
    
    If (g_paCursorPos.Y < m_MyPosition.Bottom) And _
            (g_paCursorPos.Y > m_MyPosition.Top + Me.ScaleHeight - 30) Then

        IsCursorOnBoarder = True
        Exit Function
    End If
    
    If (g_paCursorPos.X > m_MyPosition.Left) And _
            (g_paCursorPos.X < m_MyPosition.Left + 22) Then

        IsCursorOnBoarder = True
        Exit Function
    End If
    
    If (g_paCursorPos.Y > (m_MyPosition.Top + 10)) And _
            (g_paCursorPos.Y < (m_MyPosition.Top + 15)) Then

        IsCursorOnBoarder = True
        Exit Function
    End If

End Function

Public Property Get LayeredWindow() As LayerdWindowHandles
    Set LayeredWindow = m_layeredWindow
End Property

Public Property Get Graphics() As GDIPGraphics
    Set Graphics = m_graphics
End Property

Private Function DrawSlices()
    On Error GoTo Handler

Dim thisSlice As Slice

    For Each thisSlice In m_slices
        Select Case thisSlice.Anchor
        
        Case AnchorPointConstants.apTopLeft
            If thisSlice.StretchX And Not thisSlice.StretchY Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, Me.ScaleWidth - thisSlice.PixelGap, thisSlice.Height
            ElseIf thisSlice.StretchY And Not thisSlice.StretchX Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, thisSlice.Width, Me.ScaleHeight - thisSlice.PixelGap
            ElseIf thisSlice.StretchX And thisSlice.StretchY Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, Me.ScaleWidth - thisSlice.PixelGap, Me.ScaleHeight - thisSlice.PixelGap2
            Else
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, thisSlice.Width, thisSlice.Height
            End If
        
        Case AnchorPointConstants.apTopRight
            If thisSlice.StretchY Then
                m_graphics.DrawImage thisSlice.Image, Me.ScaleWidth - thisSlice.X, thisSlice.Y, thisSlice.Width, Me.ScaleHeight - thisSlice.PixelGap
            Else
                m_graphics.DrawImage thisSlice.Image, Me.ScaleWidth - thisSlice.X, thisSlice.Y, thisSlice.Width, thisSlice.Height
            End If
        
        Case AnchorPointConstants.apBottomLeft
            If thisSlice.StretchX Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, Me.ScaleHeight - thisSlice.Y, Me.ScaleWidth - thisSlice.PixelGap, thisSlice.Height
            Else
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, Me.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height
            End If
        
        Case AnchorPointConstants.apBottomRight
            m_graphics.DrawImage thisSlice.Image, Me.ScaleWidth - thisSlice.X, Me.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height

        End Select
    Next
    
    
    Exit Function
Handler:
End Function

Private Function CreateSlices()

Dim slicesXMlDocument As New DOMDocument
Dim thisXMLSlice As IXMLDOMElement
Dim thisSlice As Slice

Dim sliceWidth As Long
Dim sliceHeight As Long

    If slicesXMlDocument.Load(App.Path & "\resources\glass_index.xml") = False Then
        Exit Function
    End If
    
    Set backgroundImage = New GDIPImage
    backgroundImage.FromFile App.Path & "\Resources\glass.png"
    
    For Each thisXMLSlice In slicesXMlDocument.firstChild.childNodes
        If thisXMLSlice.tagName = "slice" Then
            
            Set thisSlice = New Slice
            
            thisSlice.Anchor = AnchorPointTextToLong(thisXMLSlice.getAttribute("anchor"))
            
            thisSlice.X = CLng(thisXMLSlice.getAttribute("x"))
            thisSlice.Y = CLng(thisXMLSlice.getAttribute("y"))
            
            sliceWidth = CLng(thisXMLSlice.getAttribute("width"))
            sliceHeight = CLng(thisXMLSlice.getAttribute("height"))
            
            If thisSlice.Anchor = apTopRight Then
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                backgroundImage.Width - thisSlice.X, _
                                                                thisSlice.Y))
            ElseIf thisSlice.Anchor = apTopLeft Then
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                thisSlice.X, _
                                                                thisSlice.Y))
            ElseIf thisSlice.Anchor = apTop Then
            
                thisSlice.Anchor = apTopLeft
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                thisSlice.X, _
                                                                thisSlice.Y))
                
                thisSlice.StretchX = True
                thisSlice.PixelGap = backgroundImage.Width - sliceWidth
                
            ElseIf thisSlice.Anchor = apLeft Then
            
                thisSlice.Anchor = apTopLeft
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                thisSlice.X, _
                                                                thisSlice.Y))
                
                thisSlice.StretchY = True
                thisSlice.PixelGap = backgroundImage.Height - sliceHeight
                
            ElseIf thisSlice.Anchor = apBottomLeft Then
                
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                0, _
                                                                backgroundImage.Height - thisSlice.Y))
            ElseIf thisSlice.Anchor = apBottomRight Then
            
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                backgroundImage.Width - thisSlice.X, _
                                                                backgroundImage.Height - thisSlice.Y))
            
            ElseIf thisSlice.Anchor = apBottom Then
                thisSlice.Anchor = apBottomLeft
                
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                thisSlice.X, _
                                                                backgroundImage.Height - thisSlice.Y))
                                                       
                thisSlice.StretchX = True
                thisSlice.PixelGap = backgroundImage.Width - sliceWidth
                
            ElseIf thisSlice.Anchor = apRight Then
            
                thisSlice.Anchor = apTopRight
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                backgroundImage.Width - thisSlice.X, _
                                                                thisSlice.Y))
                
                thisSlice.StretchY = True
                thisSlice.PixelGap = backgroundImage.Height - sliceHeight
            
            ElseIf thisSlice.Anchor = apMiddle Then
            
                Set m_clientSlice = thisSlice
            
                thisSlice.Anchor = apTopLeft
                Set thisSlice.Image = _
                    CreateNewImageFromSection(backgroundImage, _
                                                    CreateRectL(sliceHeight, _
                                                                sliceWidth, _
                                                                thisSlice.X, _
                                                                thisSlice.Y))
                                                                
                thisSlice.StretchX = True
                thisSlice.StretchY = True
                
                thisSlice.PixelGap2 = backgroundImage.Height - sliceHeight
                thisSlice.PixelGap = backgroundImage.Width - sliceWidth
            End If
        
            m_slices.Add thisSlice
        End If
        
    Next
End Function

Public Function UpdateWindow()
    Debug.Print "GlassWindow::UpdateWindow()"

    If m_layeredWindow Is Nothing Then Exit Function
    m_layeredWindow.Update Me.hWnd, Me.hdc
End Function

Private Function ResizeWindow()
    On Error GoTo Handler
    
    If m_layeredWindow Is Nothing Then Exit Function

    Debug.Print "ResizeWindow::UpdateWindow()"

    If Me Is Nothing Then Exit Function

    Set m_layeredWindow = Nothing
    Set m_layeredWindow = MakeLayerdWindow(Me)
    
    'UpdateMe
    
    m_graphics.FromHDC m_layeredWindow.theDC
    m_graphics.SmoothingMode = SmoothingModeHighSpeed
    m_graphics.PixelOffsetMode = PixelOffsetModeHighSpeed
    'm_graphics.CompositingMode = CompositingModeSourceCopy
    m_graphics.CompositingQuality = CompositingQualityHighSpeed
    m_graphics.InterpolationMode = InterpolationModeNearestNeighbor
    
Handler:
End Function

Public Function DrawGlass(ByRef Graphics As GDIPGraphics)
    Set m_graphics = Graphics
    
    If Me Is Nothing Then Exit Function
    

    m_graphics.Clear
    'm_graphics.DrawImage backgroundImage, 0, 0, backgroundImage.Width, backgroundImage.Height
    DrawSlices
    
    'm_layeredWindow.Update me.hWnd, me.hdc
    'Call UpdateLayeredWindow(me.hWnd, me.hDC, ByVal 0&, m_layeredWindow.GetSize, m_layeredWindow.theDC, m_layeredWindow.GetPoint, 0, m_layeredWindow.GetBlend, ULW_ALPHA)
End Function

Public Function Update()
    If Me Is Nothing Then Exit Function
    
    If m_registeredWidth <> Me.ScaleWidth Or m_registeredHeight <> Me.ScaleHeight Then
        Debug.Print "Re-Creating Graphics!"
        
        m_registeredWidth = Me.ScaleWidth
        m_registeredHeight = Me.ScaleHeight
        
        Set m_layeredWindow = Nothing
        Set m_layeredWindow = MakeLayerdWindow(Me)
        
        m_graphics.FromHDC m_layeredWindow.theDC
    End If
    
    m_graphics.Clear
    'm_graphics.DrawImage backgroundImage, 0, 0, backgroundImage.Width, backgroundImage.Height
    DrawSlices
    
    m_layeredWindow.Update Me.hWnd, Me.hdc
    
    'Call UpdateLayeredWindow(me.hWnd, me.hDC, ByVal 0&, m_layeredWindow.GetSize, m_layeredWindow.theDC, m_layeredWindow.GetPoint, 0, m_layeredWindow.GetBlend, ULW_ALPHA)
End Function

Public Function AttachForm(ByRef frmSource As Form)
    Set m_client = frmSource
    m_oldParenthWnd = SetWindowLong(frmSource.hWnd, GWL_HWNDPARENT, Me.hWnd)
    
    'frmSource.Hide
    'frmSource.Show vbModeless, Me
    
    'SetParent frmSource.hWnd, Me.hWnd
    
    
    'UpdateClientPosition
    
    Timer1.Enabled = True
End Function

Private Sub Form_DblClick()
    RaiseEvent onDblClick
End Sub

Private Sub Form_Initialize()
    InitializeGDIIfNotInitialized
    
    Set m_graphics = New GDIPGraphics
    'Set m_glassWindow = New GlassWindow
    Set m_slices = New Collection
    
    m_hasFocus = True

    Set m_layeredWindow = MakeLayerdWindow(Me)
    CreateSlices
    
    Call HookWindow(Me.hWnd, Me)
    
    If Config Is Nothing Then Exit Sub
    If Config.TopMostWindow Then StayOnTop Me, True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    RaiseEvent onKeyPress(CLng(KeyAscii))
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "APPICON", True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'RaiseEvent onMouseMove(Button, X - m_clientSlice.X, Y - m_clientSlice.Y)
    RaiseEvent onMouseDown(Button, X - m_clientSlice.X, Y - m_clientSlice.Y)
End Sub

Private Sub Form_MouseLeave()
    If m_hasFocus Then
        Screen.MousePointer = 4
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim onItem As Boolean
    RaiseEvent QueryOnItem(X - m_clientSlice.X, Y - m_clientSlice.Y, onItem)
    
    If onItem Then
        RaiseEvent onMouseMove(Button, X - m_clientSlice.X, Y - m_clientSlice.Y)
        Exit Sub
    End If

Dim bRight As Boolean
Dim bBottom As Boolean
Dim bLeft As Boolean
Dim bTop As Boolean

Dim bChanged As Boolean

Dim g_paCursorPos As win.POINTL
Dim m_MyPosition As win.RECT

    If m_trackingMouse = False Then
        m_trackingMouse = TrackMouse(Me.hWnd)
    End If

    GetCursorPos g_paCursorPos
    GetWindowRect Me.hWnd, m_MyPosition
    
    m_POS = HTCAPTION
    
        If m_hasFocus Then
        
            'Border Checks
            If (g_paCursorPos.X < m_MyPosition.Right) And _
                (g_paCursorPos.X > m_MyPosition.Left + Me.ScaleWidth - 22) Then
                bRight = True
            End If
            
            If (g_paCursorPos.Y < m_MyPosition.Bottom) And _
                    (g_paCursorPos.Y > m_MyPosition.Top + Me.ScaleHeight - 30) Then
                bBottom = True
            End If
            
            If (g_paCursorPos.X > m_MyPosition.Left) And _
                    (g_paCursorPos.X < m_MyPosition.Left + 22) Then
                bLeft = True
            End If
            
            If (g_paCursorPos.Y > (m_MyPosition.Top + 10)) And _
                    (g_paCursorPos.Y < (m_MyPosition.Top + 15)) Then
                bTop = True
            End If
        
            If Me.WindowState = vbNormal Then
                If (bBottom And bRight) Then
                    Screen.MousePointer = vbSizeNWSE
                    m_POS = HTBOTTOMRIGHT
                ElseIf (bBottom And bLeft) Then
                    Screen.MousePointer = vbSizeNESW
                    m_POS = HTBOTTOMLEFT
                ElseIf (bTop And bLeft) Then
                    Screen.MousePointer = vbSizeNWSE
                    m_POS = HTTOPLEFT
                ElseIf (bTop And bRight) Then
                    Screen.MousePointer = vbSizeNESW
                    m_POS = HTTOPRIGHT
                ElseIf bRight Then
                    Screen.MousePointer = vbSizeWE
                    m_POS = HTRIGHT
                ElseIf bBottom Then
                    Screen.MousePointer = vbSizeNS
                    m_POS = HTBOTTOM
                ElseIf bLeft Then
                    Screen.MousePointer = vbSizeWE
                    m_POS = HTLEFT
                ElseIf bTop Then
                    Screen.MousePointer = vbSizeNS
                    m_POS = HTTOP
                Else
                    'If IsMouseInRect(m_recTitleBar) Then
                        'm_bMouseInTitleBar = True
                    'Else
                        'm_bMouseInTitleBar = False
                    'End If
                
                    m_POS = HTCAPTION
                    Screen.MousePointer = 4
                    
                    'Event_MouseMove = True
                End If
            End If
        End If

    If bChanged Then
        If m_hasFocus Then
            Screen.MousePointer = 4
        End If
    
        'Me.ReRender
        'Me.RenderBackbufferDC
    End If
    
    If Button = vbLeftButton Then
        If m_POS > 0 And Me.WindowState = vbNormal Then
            'Tricks to resize the window
            
            Debug.Print "Just here pal!"
            
            UpdateClientPosition
            
            SetWindowLong Me.hWnd, GWL_STYLE, WS_VISIBLE Or WS_MINIMIZEBOX
   
            ReleaseCapture
            Call SendMessage(ByVal Me.hWnd, ByVal WM_NCLBUTTONDOWN, ByVal m_POS, 0&)

            'Tricks to resize the window
            SetWindowLong Me.hWnd, GWL_STYLE, WS_VISIBLE Or WS_MINIMIZEBOX Or WS_SYSMENU
            
            UpdateClientPosition
        End If
    End If
    
    If m_POS = HTCAPTION Then
        RaiseEvent onMouseMove(Button, X - m_clientSlice.X, Y - m_clientSlice.Y)
    End If

End Sub

Sub UpdateClientPosition()
    On Error GoTo Handler
    
    m_clientSize.Width = Me.ScaleWidth - m_clientSlice.PixelGap
    m_clientSize.Height = Me.ScaleHeight - m_clientSlice.PixelGap2
    
    If Not m_client Is Nothing Then
        If m_client.WindowState = vbNormal Then m_client.Move Me.Left + (m_clientSlice.X * Screen.TwipsPerPixelX), Me.Top + (m_clientSlice.Y * Screen.TwipsPerPixelX), m_clientSize.Width * Screen.TwipsPerPixelX, m_clientSize.Height * Screen.TwipsPerPixelY
    End If
    
Handler:
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent onMouseUp(Button, X - m_clientSlice.X, Y - m_clientSlice.Y)
End Sub

Private Sub Form_Resize()
    If m_invalid Then Exit Sub

    ResizeWindow
        
    DrawGlass m_graphics
    UpdateWindow
    
    UpdateClientPosition
    
    RaiseEvent onResized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not AppZorderKeeper Is Nothing Then AppZorderKeeper.RemoveChild Me
    
    UnMakeLayeredWindow Me, m_layeredWindow
    m_layeredWindow.Release
    
    Call UnhookWindow(Me.hWnd)
    
    If m_invalid Then Exit Sub
    If Not m_client Is Nothing Then
        'Unload m_client
    End If
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
On Error GoTo Handler

Dim bEat As Boolean

    If m_invalid Then GoTo Handler

    If msg = WM_MOVE Then
        UpdateClientPosition
    ElseIf msg = WM_MOUSELEAVE Then
        m_trackingMouse = False
        Form_MouseLeave
    ElseIf msg = WM_GETMINMAXINFO Then
        'Dont tell VB about this Windows Message
        bEat = True
    
        Dim MMI As MINMAXINFO
        CopyMemory MMI, ByVal lp, ByVal LenB(MMI)
            
        'set the MINMAXINFO data to the
        'minimum and maximum values set
        'by the option choice
        With MMI
            .ptMinTrackSize.X = 340
            .ptMinTrackSize.Y = 240
            
            '.ptMaxTrackSize.X = 600
            '.ptMaxTrackSize.Y = 600
        End With
      
        CopyMemory ByVal lp, MMI, ByVal LenB(MMI)
        
    ElseIf msg = WM_DROPFILES Then
        If Not m_client Is Nothing Then
            PostMessage ByVal m_client.hWnd, ByVal msg, ByVal wp, ByVal lp
        End If
    End If

Handler:
    If Not bEat Then
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wp, lp)
    End If
End Function

Private Sub Timer1_Timer()
    Me.Move m_client.Left - ((m_clientSlice.PixelGap2 / 2) * Screen.TwipsPerPixelX), _
            m_client.Top - ((m_clientSlice.PixelGap / 2) * Screen.TwipsPerPixelY), _
            m_client.Width + ((m_clientSlice.PixelGap / 2) * Screen.TwipsPerPixelX), _
            m_client.Height + ((m_clientSlice.PixelGap) * Screen.TwipsPerPixelY)
    
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    'Timer2.Enabled = False
    
    If Not m_client Is Nothing Then
        m_client.ZOrder
        m_client.Show vbModeless, Me
    End If
End Sub

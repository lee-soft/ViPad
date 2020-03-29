VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlassWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_layeredWindow As LayerdWindowHandles
Private m_form As Form

Private m_slices As Collection
Private m_graphics As GDIPGraphics

Private m_registeredWidth As Long
Private m_registeredHeight As Long

Dim backgroundImage As GDIPImage

Private m_POS As Long
Private m_hasFocus As Boolean

Public Function IsCursorOnBoarder() As Boolean

Dim g_paCursorPos As win.POINTL
Dim m_MyPosition As win.RECT

    GetCursorPos g_paCursorPos
    GetWindowRect m_form.hWnd, m_MyPosition
    
    If (g_paCursorPos.X < m_MyPosition.Right) And _
        (g_paCursorPos.X > m_MyPosition.Left + m_form.ScaleWidth - 22) Then
        
        IsCursorOnBoarder = True
        Exit Function
    End If
    
    If (g_paCursorPos.Y < m_MyPosition.Bottom) And _
            (g_paCursorPos.Y > m_MyPosition.Top + m_form.ScaleHeight - 30) Then

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

Public Function Event_MouseMove(Button As MouseButtonConstants) As Boolean

Dim bRight As Boolean
Dim bBottom As Boolean
Dim bLeft As Boolean
Dim bTop As Boolean

Dim bChanged As Boolean

Dim g_paCursorPos As win.POINTL
Dim m_MyPosition As win.RECT

    GetCursorPos g_paCursorPos
    GetWindowRect m_form.hWnd, m_MyPosition
    
    m_POS = HTCAPTION
    
        If m_hasFocus Then
        
            'Border Checks
            If (g_paCursorPos.X < m_MyPosition.Right) And _
                (g_paCursorPos.X > m_MyPosition.Left + m_form.ScaleWidth - 22) Then
                bRight = True
            End If
            
            If (g_paCursorPos.Y < m_MyPosition.Bottom) And _
                    (g_paCursorPos.Y > m_MyPosition.Top + m_form.ScaleHeight - 30) Then
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
        
            If m_form.WindowState = vbNormal Then
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
                    
                    Event_MouseMove = True
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
        If m_POS > 0 And m_form.WindowState = vbNormal Then
            'Tricks to resize the window
            
            Debug.Print "Just here pal!"
            
            SetWindowLong m_form.hWnd, GWL_STYLE, WS_VISIBLE Or WS_MINIMIZEBOX
   
            ReleaseCapture
            Call SendMessage(ByVal m_form.hWnd, ByVal WM_NCLBUTTONDOWN, ByVal m_POS, 0&)

            'Tricks to resize the window
            SetWindowLong m_form.hWnd, GWL_STYLE, WS_VISIBLE Or WS_MINIMIZEBOX Or WS_SYSMENU
        End If
    End If

End Function

Public Property Get LayeredWindow() As LayerdWindowHandles
    Set LayeredWindow = m_layeredWindow
End Property

Public Property Get Graphics() As GDIPGraphics
    Set Graphics = m_graphics
End Property

Private Function DrawSlices()

Dim thisSlice As Slice

    For Each thisSlice In m_slices
        Select Case thisSlice.Anchor
        
        Case AnchorPointConstants.apTopLeft
            If thisSlice.StretchX And Not thisSlice.StretchY Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, m_form.ScaleWidth - thisSlice.PixelGap, thisSlice.Height
            ElseIf thisSlice.StretchY And Not thisSlice.StretchX Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, thisSlice.Width, m_form.ScaleHeight - thisSlice.PixelGap
            ElseIf thisSlice.StretchX And thisSlice.StretchY Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, m_form.ScaleWidth - thisSlice.PixelGap, m_form.ScaleHeight - thisSlice.PixelGap2
            Else
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, thisSlice.Y, thisSlice.Width, thisSlice.Height
            End If
        
        Case AnchorPointConstants.apTopRight
            If thisSlice.StretchY Then
                m_graphics.DrawImage thisSlice.Image, m_form.ScaleWidth - thisSlice.X, thisSlice.Y, thisSlice.Width, m_form.ScaleHeight - thisSlice.PixelGap
            Else
                m_graphics.DrawImage thisSlice.Image, m_form.ScaleWidth - thisSlice.X, thisSlice.Y, thisSlice.Width, thisSlice.Height
            End If
        
        Case AnchorPointConstants.apBottomLeft
            If thisSlice.StretchX Then
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, m_form.ScaleHeight - thisSlice.Y, m_form.ScaleWidth - thisSlice.PixelGap, thisSlice.Height
            Else
                m_graphics.DrawImage thisSlice.Image, thisSlice.X, m_form.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height
            End If
        
        Case AnchorPointConstants.apBottomRight
            m_graphics.DrawImage thisSlice.Image, m_form.ScaleWidth - thisSlice.X, m_form.ScaleHeight - thisSlice.Y, thisSlice.Width, thisSlice.Height

        End Select
    Next
    
    

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
    backgroundImage.FromFile App.Path & "Resources\glass.png"
    
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

    m_layeredWindow.Update m_form.hWnd, m_form.hdc
End Function

Public Function ResizeWindow()
    Debug.Print "ResizeWindow::UpdateWindow()"

    If m_form Is Nothing Then Exit Function

    Set m_layeredWindow = Nothing
    Set m_layeredWindow = MakeLayerdWindow(m_form)
    
    'UpdateMe
End Function

Public Function DrawGlass(ByRef Graphics As GDIPGraphics)
    Set m_graphics = Graphics
    
    If m_form Is Nothing Then Exit Function
    

    m_graphics.Clear
    'm_graphics.DrawImage backgroundImage, 0, 0, backgroundImage.Width, backgroundImage.Height
    DrawSlices
    
    'm_layeredWindow.Update m_form.hWnd, m_form.hdc
    'Call UpdateLayeredWindow(m_form.hWnd, m_form.hDC, ByVal 0&, m_layeredWindow.GetSize, m_layeredWindow.theDC, m_layeredWindow.GetPoint, 0, m_layeredWindow.GetBlend, ULW_ALPHA)
End Function

Public Function Update()
    If m_form Is Nothing Then Exit Function
    
    If m_registeredWidth <> m_form.ScaleWidth Or m_registeredHeight <> m_form.ScaleHeight Then
        Debug.Print "Re-Creating Graphics!"
        
        m_registeredWidth = m_form.ScaleWidth
        m_registeredHeight = m_form.ScaleHeight
        
        Set m_layeredWindow = Nothing
        Set m_layeredWindow = MakeLayerdWindow(m_form)
        
        m_graphics.FromHDC m_layeredWindow.theDC
    End If
    
    m_graphics.Clear
    'm_graphics.DrawImage backgroundImage, 0, 0, backgroundImage.Width, backgroundImage.Height
    DrawSlices
    
    m_layeredWindow.Update m_form.hWnd, m_form.hdc
    'Call UpdateLayeredWindow(m_form.hWnd, m_form.hDC, ByVal 0&, m_layeredWindow.GetSize, m_layeredWindow.theDC, m_layeredWindow.GetPoint, 0, m_layeredWindow.GetBlend, ULW_ALPHA)
End Function

Public Function AttachForm(ByRef frmSource As Form)
    Set m_form = frmSource
    Set m_layeredWindow = MakeLayerdWindow(m_form)

    CreateSlices
End Function

Private Sub Class_Initialize()
    m_hasFocus = True

    Set m_slices = New Collection
    'Set m_Graphics = New GDIPGraphics
End Sub


VERSION 5.00
Begin VB.Form OptionsWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timActivateDelay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "OptionsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WINDOWS_REGRUN As String = "Software\Microsoft\Windows\CurrentVersion\Run\"

Public Event onNewIconSize(newSize As Long)
Public Event onClose()
Public Event onStickToDesktop()
Public Event onLayeredModeChanged()
Public Event onWindowStyleChanged()

Private WithEvents m_cmdClose As ViCommandButton
Attribute m_cmdClose.VB_VarHelpID = -1
Private WithEvents m_cmdOK As ViCommandButton
Attribute m_cmdOK.VB_VarHelpID = -1
Private WithEvents m_IconSize As ViSlider
Attribute m_IconSize.VB_VarHelpID = -1

Private WithEvents m_chkStartWithWindows As ViCheckBox
Attribute m_chkStartWithWindows.VB_VarHelpID = -1
Private WithEvents m_chkStickToDesktop As ViCheckBox
Attribute m_chkStickToDesktop.VB_VarHelpID = -1
Private WithEvents m_chkTopMost As ViCheckBox
Attribute m_chkTopMost.VB_VarHelpID = -1
Private WithEvents m_chkMinimizeAfterLaunch As ViCheckBox
Attribute m_chkMinimizeAfterLaunch.VB_VarHelpID = -1
Private WithEvents m_chkHideDesktop As ViCheckBox
Attribute m_chkHideDesktop.VB_VarHelpID = -1
Private WithEvents m_chkDWMMode As ViCheckBox
Attribute m_chkDWMMode.VB_VarHelpID = -1
Private WithEvents m_chkControlBox As ViCheckBox
Attribute m_chkControlBox.VB_VarHelpID = -1
Private WithEvents m_chkTabMode As ViCheckBox
Attribute m_chkTabMode.VB_VarHelpID = -1
Private WithEvents m_chkMiddleMouseActivator As ViCheckBox
Attribute m_chkMiddleMouseActivator.VB_VarHelpID = -1

Private m_lblIconDimensions As ViText
Private m_lblStartWithWindows As ViText
Private m_lblStickToDesktop As ViText
Private m_lblAlwaysOnTop As ViText
Private m_lblIconSize As ViText
Private m_lblInstanceMode As ViText

Private m_textLayer As ViTextLayer

Private m_graphicsWindow As GDIPGraphics
Private m_backBuffer As GDIPBitmap
Private m_glassMode As Boolean
Private m_acceptInput As Boolean

Private m_activeObject As Object
'Private m_glassText As VistaGlassHandles

Private m_viComponents As Collection

Private Sub Command1_Click()
    ReDrawComponents
    Me.Refresh
End Sub

Private Sub Form_Activate()
    timActivateDelay.Enabled = True
End Sub

Private Sub m_chkControlBox_onChange()
    If Config.ShowControlBox <> (m_chkControlBox.Checked) Then
        Config.ShowControlBox = m_chkControlBox.Checked
        
        RaiseEvent onWindowStyleChanged
    End If
End Sub

Private Sub m_chkDWMMode_onChange()
    If Config.ForceLayeredMode <> (Not m_chkDWMMode.Checked) Then
        Config.ForceLayeredMode = Not m_chkDWMMode.Checked
        RaiseEvent onLayeredModeChanged
    End If
End Sub

Private Sub m_chkHideDesktop_onChange()
    Config.HideDesktopOnBoot = m_chkHideDesktop.Checked

    If m_chkHideDesktop.Checked Then
        g_hiddenDesktop = True
        HideDesktopIcons
    Else
        If g_hiddenDesktop Then
            g_hiddenDesktop = False
            ShowDesktopIcons
            
        End If
    End If
End Sub

Private Sub m_chkMiddleMouseActivator_onChange()
Dim m_CatchingMiddleMouseClicks As Boolean
    m_CatchingMiddleMouseClicks = PluginExists("VIPAD_MIDDLE_MOUSE_NOTIFIER")
    
    Config.MiddleMouseActivation = m_chkMiddleMouseActivator.Checked
    
    If Config.MiddleMouseActivation Then
        m_chkStickToDesktop.Checked = False
    
        If Not m_CatchingMiddleMouseClicks Then
            StartPlugin "VIPAD_MIDDLE_MOUSE_NOTIFIER"
        End If
    Else
        If m_CatchingMiddleMouseClicks Then
            TerminatePlugin "VIPAD_MIDDLE_MOUSE_NOTIFIER"
        End If
    End If
End Sub

Private Sub m_chkMinimizeAfterLaunch_onChange()
    Config.MinimizeAfterLauch = m_chkMinimizeAfterLaunch.Checked
    If m_chkMinimizeAfterLaunch.Checked = False Then m_chkTopMost.Checked = False
    If m_chkMinimizeAfterLaunch.Checked Then m_chkStickToDesktop.Checked = False
End Sub

Private Sub m_chkStartWithWindows_onChange()
    If m_chkStartWithWindows.Checked Then
        WriteReg HKEY_CURRENT_USER, WINDOWS_REGRUN, App.ProductName, App.Path & "\" & App.EXEName & ".exe"
    Else
        DeleteReg HKEY_CURRENT_USER, WINDOWS_REGRUN, App.ProductName
    End If
    
End Sub

Private Sub m_chkStickToDesktop_onChange()
    Config.StickToDesktop = m_chkStickToDesktop.Checked
    
    If m_chkStickToDesktop.Checked Then
        m_chkMinimizeAfterLaunch.Checked = False
        m_chkMiddleMouseActivator.Checked = False
    End If
    
    RaiseEvent onStickToDesktop
End Sub

Private Sub m_chkTabMode_onChange()
    If Config.InstanceMode <> (Not m_chkTabMode.Checked) Then
        ToggleAppType
    End If
End Sub

Private Sub m_chkTopMost_onChange()

Dim thisForm As Form
    Config.TopMostWindow = m_chkTopMost.Checked
    
    If m_chkTopMost.Checked Then
        m_chkMinimizeAfterLaunch.Checked = m_chkTopMost.Checked
        m_chkStickToDesktop.Checked = False
    End If
    
    For Each thisForm In Forms
        StayOnTop thisForm, Config.TopMostWindow
    Next
End Sub

'
'
'

Private Sub Form_Load()
    InitializeGDIIfNotInitialized
    
    SetIcon Me.hWnd, "APPICON", True
    
    'Set m_glassText = InitializeFormForGlassText(Me)
    Set m_viComponents = New Collection
    
    Set m_backBuffer = New GDIPBitmap
    Set m_graphicsWindow = New GDIPGraphics
    
    'm_glassText.AttachForm Me
    'm_glassText.OpenWindowForGlass 'creates a DC for pointer hGlassDc

    Set m_IconSize = New ViSlider
    
    Set m_chkStartWithWindows = New ViCheckBox
    Set m_chkStickToDesktop = New ViCheckBox
    Set m_chkTopMost = New ViCheckBox
    Set m_chkMinimizeAfterLaunch = New ViCheckBox
    Set m_chkHideDesktop = New ViCheckBox
    Set m_chkDWMMode = New ViCheckBox
    Set m_chkControlBox = New ViCheckBox
    Set m_chkTabMode = New ViCheckBox
    Set m_chkMiddleMouseActivator = New ViCheckBox
    
    Set m_cmdClose = New ViCommandButton
    Set m_cmdOK = New ViCommandButton
    
    Set m_textLayer = New ViTextLayer
    Set m_lblIconDimensions = New ViText
    Set m_lblStartWithWindows = New ViText
    Set m_lblStickToDesktop = New ViText
    Set m_lblAlwaysOnTop = New ViText
    
    m_textLayer.Parent = Me
    
    m_chkStartWithWindows.X = 16
    m_chkStartWithWindows.Y = 88
    
    m_chkStickToDesktop.X = 16
    m_chkStickToDesktop.Y = 132
    
    m_chkTopMost.X = 16
    m_chkTopMost.Y = 176
    
    m_chkMinimizeAfterLaunch.X = 16
    m_chkMinimizeAfterLaunch.Y = 220
    
    m_chkHideDesktop.X = 16
    m_chkHideDesktop.Y = 264
    
    m_chkTabMode.X = 16
    m_chkTabMode.Y = 308
    
    m_chkMiddleMouseActivator.X = 16
    m_chkMiddleMouseActivator.Y = 352
    
    m_chkDWMMode.X = 16
    m_chkDWMMode.Y = 396
    
    m_chkControlBox.X = 16
    m_chkControlBox.Y = 440
    
    m_cmdOK.X = 204
    m_cmdOK.Y = 500
    m_cmdOK.Width = 69
    m_cmdOK.Caption = "OK"
    
    m_cmdClose.X = 16
    m_cmdClose.Y = 500
    m_cmdClose.Width = 133
    m_cmdClose.Caption = "Close Program"
    
    m_textLayer.CreateChild "Icon Dimensions", _
                            28, _
                            15, _
                            Me.FontName, _
                            Me.fontSize
                            
    Set m_lblIconSize = m_textLayer.CreateChild("0", _
                            180, _
                            15, _
                            Me.FontName, _
                            Me.fontSize)
                            
    m_textLayer.CreateChild "Start With Windows", _
                            60, _
                            97, _
                            Me.FontName, _
                            Me.fontSize
    
    m_textLayer.CreateChild "Stick to Desktop (WinKey + D)", _
                            60, _
                            141, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "Always On Top", _
                            60, _
                            185, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "Minimize After Launch", _
                            60, _
                            229, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "Hide Desktop", _
                            60, _
                            273, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "Tabbed Mode", _
                            60, _
                            317, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "Middle Mouse Dbl-Click Activation", _
                            60, _
                            363, _
                            Me.FontName, _
                            Me.fontSize - 1
                            
    
    m_IconSize.X = 16
    m_IconSize.Y = 48
    m_IconSize.Width = 257
    m_IconSize.Max = 256
    m_IconSize.Min = 32
    
    m_viComponents.Add m_IconSize
    m_viComponents.Add m_chkStartWithWindows
    m_viComponents.Add m_chkStickToDesktop
    m_viComponents.Add m_chkTopMost
    m_viComponents.Add m_chkMinimizeAfterLaunch
    m_viComponents.Add m_chkHideDesktop
    m_viComponents.Add m_chkTabMode
    m_viComponents.Add m_chkMiddleMouseActivator
    
    m_viComponents.Add m_cmdClose
    m_viComponents.Add m_cmdOK
    
    m_viComponents.Add m_textLayer
        
    If ApplyGlassIfPossible(Me.hWnd) = False Then
    Else
        m_glassMode = True
        
        m_viComponents.Add m_chkDWMMode
        m_viComponents.Add m_chkControlBox
        
        m_textLayer.CreateChild "Use Windows DWM", _
                                60, _
                                405, _
                                Me.FontName, _
                                Me.fontSize
                                
        m_textLayer.CreateChild "Show Control Box", _
                                60, _
                                449, _
                                Me.FontName, _
                                Me.fontSize
    End If
    
    ShowSettings
    If Config.TopMostWindow Then StayOnTop Me, True
    Me.Show
    
    m_graphicsWindow.FromHDC Me.hdc
    m_graphicsWindow.SmoothingMode = SmoothingModeHighQuality
    m_graphicsWindow.InterpolationMode = InterpolationModeHighQualityBicubic
    m_graphicsWindow.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    ReDrawComponents
End Sub

Private Sub ReDrawComponents()

    If m_viComponents Is Nothing Then Exit Sub

    Debug.Print "ReDrawComponents!"

    'm_glassText.OpenWindowForGlass 'Same effect as clear but doesn't crash everything
    If m_glassMode Then
        m_graphicsWindow.Clear
    Else
        m_graphicsWindow.Clear Me.BackColor
    End If
  
Dim thisObject As Object
    For Each thisObject In m_viComponents
        thisObject.Draw m_graphicsWindow
    Next
    
    'm_glassText.DrawText "Icon Dimensions", 15, 28
    'm_glassText.DrawText "Test", 1, 20
    'm_backBuffer.h

    'm_glassText.Flush
    Me.Refresh
    
    'Call RedrawWindow(Me.hWnd, ByVal 0&, 0&, _
             RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_acceptInput Then Exit Sub

Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            thisObject.onMouseDown CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        End If
    Next

    If redrawMe Then
        ReDrawComponents
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            If Not m_activeObject Is thisObject Then Set m_activeObject = thisObject
            thisObject.onMouseMove CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        Else
            If Not m_activeObject Is Nothing Then
                m_activeObject.onMouseOut
                Set m_activeObject = Nothing
            End If
        End If
    Next

    If redrawMe Then
        ReDrawComponents
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_acceptInput Then Exit Sub

Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            thisObject.onMouseUp CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        End If
    Next

    If redrawMe Then
        ReDrawComponents
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_acceptInput = False
    
    Set m_viComponents = Nothing
    Set m_textLayer = Nothing
End Sub

Private Sub m_cmdClose_onClicked()
    EndApplication
End Sub

Private Sub m_cmdOK_onClicked()
    RaiseEvent onClose
    Unload Me
End Sub

Private Sub ShowSettings()
    m_chkTopMost.Checked = Config.TopMostWindow
    m_chkHideDesktop.Checked = Config.HideDesktopOnBoot
    m_chkControlBox.Checked = Config.ShowControlBox
    m_chkMinimizeAfterLaunch.Checked = Config.MinimizeAfterLauch
    m_IconSize.Value = Config.IconSize
    m_lblIconSize.Caption = Config.IconSize & " x " & Config.IconSize
    m_chkStickToDesktop.Checked = Config.StickToDesktop
    m_chkMiddleMouseActivator.Checked = Config.MiddleMouseActivation
    
    m_chkDWMMode.Checked = Not Config.ForceLayeredMode
    m_chkTabMode.Checked = Not Config.InstanceMode
    
    If ReadReg(HKEY_CURRENT_USER, WINDOWS_REGRUN, App.ProductName) = App.Path & "\" & App.EXEName & ".exe" Then
        m_chkStartWithWindows.Checked = True
    Else
        m_chkStartWithWindows.Checked = False
    End If
End Sub

Private Sub m_IconSize_onChange(newValue As Long)
    m_lblIconSize.Caption = CStr(newValue) & " x " & CStr(newValue)
    Config.IconSize = newValue
    
    ReDrawComponents
    
    RaiseEvent onNewIconSize(newValue)
End Sub


Private Sub timActivateDelay_Timer()
    m_acceptInput = True
    timActivateDelay.Enabled = False
End Sub

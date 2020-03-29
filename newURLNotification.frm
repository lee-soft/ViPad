VERSION 5.00
Begin VB.Form URLNotification 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4995
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "URLNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_glassMode As Boolean
Private m_textLayer As ViTextLayer
'Private m_addURL As ViText
Private m_URL As ViText
Private m_URLText As String
Private m_rawURL As String
Private m_physicalShortcutFile As String

Private WithEvents m_cmdCancel As ViCommandButton
Attribute m_cmdCancel.VB_VarHelpID = -1
Private WithEvents m_httpRequest As WinHttpRequest
Attribute m_httpRequest.VB_VarHelpID = -1

Private m_viComponents As Collection
Private m_backBuffer As GDIPBitmap
Private m_graphicsWindow As GDIPGraphics
Private m_activeObject As Object
Private m_activeIcon As LaunchPadItem
Private m_requestData As String
Private m_hWndNextViewer As Long
Private m_showURLCatcher As Boolean

Public Event onLoad()
Public Event onCancel()
Public Event onConfirm()

Implements IHookSink

Public Property Get Ready() As Boolean
    Ready = m_showURLCatcher
End Property

Public Property Get Title() As String
    Title = m_URLText
End Property

Public Property Get URL() As String
    URL = m_rawURL
End Property

Public Property Let URL(newURL As String)

    m_requestData = ""
    
    m_URLText = newURL
    m_rawURL = newURL

    Debug.Print newURL

    Set m_httpRequest = New WinHttpRequest
    m_httpRequest.Open "GET", newURL, True
    m_httpRequest.Send
End Property

Private Sub Form_Initialize()
    HookWindow Me.hWnd, Me
    m_hWndNextViewer = SetClipboardViewer(Me.hWnd)
End Sub

'
'
'

Private Sub Form_Load()

Dim thisText As ViText

    InitializeGDIIfNotInitialized
    
    SetIcon Me.hWnd, "APPICON", True
    
    'Set m_glassText = InitializeFormForGlassText(Me)
    Set m_viComponents = New Collection
    
    Set m_backBuffer = New GDIPBitmap
    Set m_graphicsWindow = New GDIPGraphics
    
    Set m_cmdCancel = New ViCommandButton
    Set m_textLayer = New ViTextLayer
    Set m_activeIcon = New LaunchPadItem
    
    Set m_activeIcon = New LaunchPadItem
    m_activeIcon.X = 106
    m_activeIcon.Y = 40
    
    Set m_activeIcon.AlphaIconStore = ProgramSupport.DefaultBrowserIcon
    
    m_textLayer.Parent = Me
    
    m_viComponents.Add m_textLayer
    m_viComponents.Add m_cmdCancel

    m_cmdCancel.Y = 215
    m_cmdCancel.X = 125
    m_cmdCancel.Width = 93
    m_cmdCancel.Caption = "Close"
    
    If ApplyGlassIfPossible(Me.hWnd) = False Then
    Else
        m_glassMode = True
    End If
    
    Set thisText = m_textLayer.CreateChild("Drag the link to the desired window", _
                            9, _
                            2, _
                            Me.FontName, _
                            Me.fontSize)
                            
    thisText.Alignment = StringAlignmentCenter
    thisText.Width = Me.ScaleWidth
    thisText.Height = 30
    
    Set m_URL = m_textLayer.CreateChild(m_URLText, 70, 170, Me.FontName, 14)
    
    m_URL.Alignment = StringAlignmentCenter
    m_URL.Width = 190
    m_URL.Height = 30
    
    m_graphicsWindow.FromHDC Me.hdc
    m_graphicsWindow.SmoothingMode = SmoothingModeHighQuality
    m_graphicsWindow.InterpolationMode = InterpolationModeHighQualityBicubic
    m_graphicsWindow.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    ReDrawComponents
    
    If Config.TopMostWindow Then StayOnTop Me, True
End Sub

Private Sub ReDrawComponents()
    If m_viComponents Is Nothing Then Exit Sub

    If m_glassMode Then
        m_graphicsWindow.Clear
    Else
        m_graphicsWindow.Clear Me.BackColor
    End If
  
Dim thisObject As Object
    For Each thisObject In m_viComponents
        thisObject.Draw m_graphicsWindow
    Next
    
    If Not m_activeIcon.AlphaIconStore Is Nothing Then
        m_graphicsWindow.DrawImage m_activeIcon.AlphaIconStore.Image, m_activeIcon.X, m_activeIcon.Y, 128, 128
    End If

    Me.Refresh
End Sub

Private Sub Form_Resize()

    m_graphicsWindow.FromHDC Me.hdc
    m_graphicsWindow.SmoothingMode = SmoothingModeHighQuality
    m_graphicsWindow.InterpolationMode = InterpolationModeHighQualityBicubic
    m_graphicsWindow.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    ReDrawComponents
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wParam As Long, lParam As Long) As Long
    On Error GoTo Handler

Dim bEat As Boolean

    If msg = WM_CHANGECBCHAIN Then
    
        ' If the next viewer window is closing, repair the chain:
        m_hWndNextViewer = lParam
        If (m_hWndNextViewer <> 0) Then
           ' Otherwise if there is a next window, pass the message on:
           SendMessage m_hWndNextViewer, msg, wParam, lParam
        End If
    
    ElseIf msg = WM_ACTIVATE Then
    
        m_showURLCatcher = False
    
    ElseIf msg = WM_DRAWCLIPBOARD Then
        ' the content of the clipboard has changed.
        ' We raise a ClipboardChanged message and pass the message on:
        
        '!!RUNTIME ERROR READING CLIPBOARD (on Error should fix it)
        If LCase(Left(Clipboard.GetText, 5)) = "http:" Then
            'If Not m_urlCatcher Is Nothing Then Unload m_urlCatcher
            'Set m_urlCatcher = New URLNotification
            m_rawURL = Clipboard.GetText
            m_URLText = "Web Link.. "
            Me.URL = Clipboard.GetText
            
            If Config.CatchWebLinks Then m_showURLCatcher = True
        Else
            If m_showURLCatcher Then
                m_showURLCatcher = False
                Me.Hide
            End If
        End If
        
        If (m_hWndNextViewer <> 0) Then
           SendMessage m_hWndNextViewer, msg, wParam, lParam
        End If
    End If

    If Not bEat Then
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wParam, lParam)
    End If

    Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wParam, lParam)
End Function

Private Sub m_cmdCancel_onClicked()
    Me.Hide
End Sub

Private Sub CreateURLLink()

    If FileExists(m_physicalShortcutFile) Then
        DeleteFile m_physicalShortcutFile
        m_physicalShortcutFile = vbNullString
    End If
    
   ' MsgBox m_URLText
    If m_rawURL = vbNullString Then Exit Sub
    If m_URLText = vbNullString Then Exit Sub
    
Dim FSO As New FileSystemObject
Dim tempLink As New ShellLinkClass
Dim urlData As String

    urlData = "[InternetShortcut]" & vbCrLf & _
              "URL=" & m_rawURL & vbCrLf

    m_physicalShortcutFile = VIPAD_SPECIAL_URL
    FSO.CreateTextFile(m_physicalShortcutFile, True, False).Write urlData
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL
Dim hitComponent As Boolean

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            hitComponent = True
            thisObject.onMouseDown CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        End If
    Next
    
    If Not hitComponent Then
        CreateURLLink
        Me.OLEDrag
    End If

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
    UnhookWindow Me.hWnd
    ChangeClipboardChain Me.hWnd, m_hWndNextViewer

    If FileExists(m_physicalShortcutFile) Then
        DeleteFile m_physicalShortcutFile
        m_physicalShortcutFile = vbNullString
    End If
    
    Set m_viComponents = Nothing
    Set m_textLayer = Nothing
End Sub

Private Sub m_cmdNo_onClicked()
    RaiseEvent onCancel
End Sub

Private Sub m_cmdYes_onClicked()
    RaiseEvent onConfirm
End Sub

Private Sub m_httpRequest_OnResponseDataAvailable(Data() As Byte)
    m_requestData = m_requestData & StrConv(Data, vbUnicode)
    If InStr(m_requestData, "</title>") > 0 Then
        m_httpRequest.Abort
        
        m_URLText = GetWebpageTitle(m_requestData)
        
        If Not m_URL Is Nothing Then
            m_URL.Caption = m_URLText
            ReDrawComponents
        End If
    End If
End Sub

Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
    Data.Files.Add m_physicalShortcutFile
    
End Sub

Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    On Error GoTo Handler
    ' Set data format to file.
    Data.SetData , vbCFFiles
    ' Display the move mouse pointer..
    AllowedEffects = vbDropEffectCopy
    
    Data.Files.Add m_physicalShortcutFile
Handler:
End Sub

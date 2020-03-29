VERSION 5.00
Begin VB.Form TrayIcon 
   BorderStyle     =   0  'None
   Caption         =   "VIPAD27081987"
   ClientHeight    =   3030
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Option Explicit

Implements IHookSink

Private m_nid As NOTIFYICONDATA
Private m_addedToSystemTray As Boolean
Private WithEvents m_systemEvents As SystemNotificationManager
Attribute m_systemEvents.VB_VarHelpID = -1

Public Event onMouseDown(Button As Long)
Public Event onMouseUp(Button As Long)
Public Event onDoubleClick(Button As Long)
Public Event onAddItem(newItemPath As String)

Public Function Notify_AddNewLink(ByVal newLnkFile As String)
    RaiseEvent onAddItem(newLnkFile)
End Function

Public Function ActivateAllWindows()
    'Triggers ViPad Windows to activate
    RaiseEvent onMouseUp(vbLeftButton)
End Function

Private Sub CopyByteArray(sourceString As String, ByRef destinationArray() As Byte)

Dim byteIndex As Long

    For byteIndex = 1 To LenB(sourceString)
        destinationArray(byteIndex - 1) = AscB(MidB(sourceString, byteIndex, 1))
    Next

End Sub


Private Sub Form_Initialize()
    With m_nid
        .cbSize = Len(m_nid)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .HICON = GetIconFromResource("APPICON")
        CopyByteArray App.ProductName, .szTip
    End With
    
    If Config.ShowTrayIcon Then
        AddToSystemTray
    End If
    
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Caption = g_APPID
    Me.Hide
    
    HookWindow Me.hWnd, Me
    
    'Subscribe to system events
    Set m_systemEvents = MainMod.SystemNotifications
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
    
    Case WM_LBUTTONUP        '514 restore form window
        'Me.WindowState = vbNormal
        'Result = SetForegroundWindow(Me.hWnd)
        'Me.Show
        RaiseEvent onMouseUp(vbLeftButton)
    
    Case WM_LBUTTONDBLCLK    '515 restore form window
        'Me.WindowState = vbNormal
        'Result = SetForegroundWindow(Me.hWnd)
        'Me.Show
        RaiseEvent onDoubleClick(vbLeftButton)
    
    Case WM_RBUTTONUP        '517 display popup menu
        'Result = SetForegroundWindow(Me.hWnd)
        'Me.PopupMenu Me.mPopupSys
        RaiseEvent onMouseUp(vbRightButton)
    
    Case Else
        Debug.Print msg
        
    End Select
End Sub

Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
    
    RemoveFromSystemTray
End Sub

Public Function AddToSystemTray()
    If Not m_addedToSystemTray Then
        Shell_NotifyIcon NIM_ADD, m_nid
        m_addedToSystemTray = True
    End If
End Function

Public Function RemoveFromSystemTray()
    If m_addedToSystemTray Then
        Shell_NotifyIcon NIM_DELETE, m_nid
        m_addedToSystemTray = False
    End If
End Function

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
On Error GoTo Handler

Dim tCDS As COPYDATASTRUCT
Dim b() As Byte

    If msg = WM_COPYDATA Then

        CopyMemory tCDS, ByVal lp, Len(tCDS)
        ReDim b(0 To tCDS.cbData) As Byte
        
        If tCDS.dwData = 88 Then
            CopyMemory b(0), ByVal tCDS.lpData, tCDS.cbData
            RaiseEvent onAddItem(CStr(b))
            
        ElseIf tCDS.dwData = 87 Then
            CopyMemory b(0), ByVal tCDS.lpData, tCDS.cbData
            RecieveAppMessage wp, CStr(b)
            
        End If
    End If

Handler:
    ' Just allow default processing for everything else.
IHookSink_WindowProc = _
   InvokeWindowProc(hWnd, msg, wp, lp)
End Function

Private Sub m_systemEvents_EndSession(ByVal EndingInitiated As Boolean, ByVal Flag As EndSessionFlags)
    EndApplication
End Sub

Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    Unload Me
End Sub

Private Sub mPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
End Sub

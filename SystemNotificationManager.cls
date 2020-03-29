VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemNotificationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IHookSink

' Member variables
Private m_hWndHidden As Long

Public Event QueryEndSession(ByVal Flag As EndSessionFlags, Cancel As Boolean)
Public Event EndSession(ByVal EndingInitiated As Boolean, ByVal Flag As EndSessionFlags)

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Hook into hidden toplevel message stream.
   m_hWndHidden = FindHiddenTopWindow()
   Call HookWindow(m_hWndHidden, Me)
End Sub

Private Sub Class_Terminate()
    UnhookWindow m_hWndHidden
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wParam As Long, LParam As Long) As Long

Dim Cancel As Boolean
Dim EatIt As Boolean

    Select Case msg
    
      Case WM_ENDSESSION
         ' wParam confirms end, lParam provides reason flag.
         ' If the session is being ended, wParam is TRUE; the session
         ' can end any time after all applications have returned from
         ' processing this message. Otherwise, it is FALSE.
         RaiseEvent EndSession(CBool(wParam), LParam)
         
      Case WM_QUERYENDSESSION
         ' lParam provides flag that indicates reason for ending.
         RaiseEvent QueryEndSession(LParam, Cancel)
         ' Applications should respect the user's intentions and
         ' return TRUE. By default, the DefWindowProc function
         ' returns TRUE for this message.
         ' If shutting down would corrupt the system or media that
         ' is being burned, the application can return FALSE.
         IHookSink_WindowProc = Abs(Not Cancel)
         EatIt = True
         
    End Select
    
    If Not EatIt Then
        ' Just allow default processing for everything else.
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, msg, wParam, LParam)
    End If

End Function

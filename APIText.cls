VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowExW Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private m_hWnd As Long
Private m_hDC As Long
Private m_font As GDIFont

Implements IHookSink

Public Event onClose()

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let ParentHwnd(newParentHwnd As Long)
    Initialize newParentHwnd
End Property

Public Property Let Text(newText As String)

Dim bString() As Byte

    bString = newText
    SendMessageW m_hWnd, WM_SETTEXT, 0&, StrPtr(bString)
    
End Property

Public Property Get Text() As String

Dim iLength As Long
Dim bString As String

    iLength = SendMessageW(m_hWnd, WM_GETTEXTLENGTH, 0&, 0&)
    
    bString = Space(iLength)
    SendMessageW m_hWnd, WM_GETTEXT, iLength + 1, StrPtr(bString)

    Text = bString

End Property

Public Property Let FontAPI(ByRef newFont As GDIFont)
    Set m_font = newFont

    SendMessage m_hWnd, WM_SETFONT, ByVal m_font.Handle, MAKELPARAM(0, False)
End Property

Private Sub Class_Terminate()
    UnhookWindow m_hWnd
    DestroyWindow m_hWnd
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

On Error GoTo Handler

    If msg = WM_CHAR Then
        
        If wp = vbKeyReturn Then
            RaiseEvent onClose
            Exit Function
        End If
    End If
    
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, msg, wp, lp)
    Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, msg, wp, lp)
End Function

Private Sub Initialize(ParentHwnd As Long)
    m_hWnd = CreateWindowExW(0, StrPtr("EDIT"), StrPtr("ViPick_Edit"), _
                WS_VISIBLE Or WS_CHILD Or WS_EX_TOOLWINDOW Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL Or ES_CENTER, _
                0, 0, 0, 0, ParentHwnd, 0, GetModuleHandle(0), 0)

Dim thisFont As New GDIFont

    thisFont.Constructor "Tahoma", 30, APIFALSE
    
    Me.FontAPI = thisFont
    
    HookWindow m_hWnd, Me
End Sub

Sub Resize(ByVal newWidth As Long, ByVal newHeight As Long)
    MoveWindow m_hWnd, 0, 0, newWidth, newHeight, True
End Sub

Private Sub Form_Terminate()
    UnhookWindow m_hWnd
    DestroyWindow m_hWnd
End Sub




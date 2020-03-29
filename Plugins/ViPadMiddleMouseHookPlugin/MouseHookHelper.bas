Attribute VB_Name = "MouseHookHelper"
Option Explicit

Private m_lowlvlMouseHook As Long
Private m_firstClickTime As Long
Private m_doubleClickTime As Long
Private m_formToActivate As frmMain

Private Const HC_ACTION = 0
Private Const WH_MOUSE_LL = 14

Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

'http://msdn.microsoft.com/en-us/library/windows/desktop/ms644974%28v=vs.85%29.aspx
Private Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, LParam As Any) As Long

'http://msdn.microsoft.com/en-us/library/windows/desktop/ms644990%28v=vs.85%29.aspx
Private Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

'http://msdn.microsoft.com/en-us/library/windows/desktop/ms644993%28v=vs.85%29.aspx
Private Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long

Private Declare Function GetDoubleClickTime Lib "user32.dll" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MSLLHOOKSTRUCT 'Will Hold the lParam struct Data
    pt As POINTAPI
    mouseData As Long ' Holds Forward\Bacward flag
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Function InstallHook(ByRef theFormToActivate As Form)
    RemoveHook
    
    Set m_formToActivate = theFormToActivate
    m_lowlvlMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LowLevelMouseProc, App.hInstance, 0)
    m_doubleClickTime = GetDoubleClickTime()
End Function

Public Sub RemoveHook()
    If m_lowlvlMouseHook = 0 Then Exit Sub
    UnhookWindowsHookEx m_lowlvlMouseHook
End Sub

Public Function LowLevelMouseProc _
    (ByVal nCode As Long, ByVal wParam As Long, ByRef LParam As MSLLHOOKSTRUCT) As Long

    If (nCode = HC_ACTION) Then
        If wParam = WM_MBUTTONUP Then

            If (m_firstClickTime + m_doubleClickTime) > GetTickCount() Then
                m_formToActivate.NotifyMiddleMouseClick
            End If
            
            m_firstClickTime = GetTickCount()
        End If

    End If

    LowLevelMouseProc = CallNextHookEx(m_lowlvlMouseHook, nCode, wParam, LParam)
End Function

Public Function SetKeyboardActiveWindow(hWnd As Long) As Boolean

Dim ForegroundThreadID As Long
Dim ThisThreadID As Long

    SetKeyboardActiveWindow = False

    ForegroundThreadID = GetWindowThreadProcessId(GetForegroundWindow(), 0&)
    ThisThreadID = GetWindowThreadProcessId(hWnd, 0&)
    
    If AttachThreadInput(ThisThreadID, ForegroundThreadID, 1) = 1 Then
        BringWindowToTop hWnd
        SetForegroundWindow hWnd
        AttachThreadInput ThisThreadID, ForegroundThreadID, 0
        SetKeyboardActiveWindow = (GetForegroundWindow = hWnd)
    End If

End Function

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Sub Main()
    If Not App.PrevInstance Then Load frmMain
End Sub

Attribute VB_Name = "TaskbarHelper"
Option Explicit

Public g_ReBarWindow32Hwnd As Long
Public g_RunningProgramsHwnd As Long
Public g_StartButtonHwnd As Long
Public g_TaskBarHwnd As Long
Public g_StartMenuHwnd As Long
Public g_StartMenuOpen As Boolean
Public g_viStartRunning As Boolean
Public g_viStartOrbHwnd As Long

Public g_WindowsVista As Boolean

Private m_TargetTaskList As TaskList
Private m_taskbarRect As Win.RECT

Public Function UpdatehWnds() As Boolean
Dim newTaskBarHwnd As Long
Dim updatedHwnd As Boolean

    updatedHwnd = False
    newTaskBarHwnd = FindWindow("Shell_TrayWnd", "")
    
    If newTaskBarHwnd = 0 Then
        Exit Function
    End If

    If newTaskBarHwnd <> g_TaskBarHwnd Then
        g_TaskBarHwnd = newTaskBarHwnd
        g_ReBarWindow32Hwnd = FindWindowEx(ByVal g_TaskBarHwnd, ByVal 0&, "ReBarWindow32", vbNullString)
        'g_ReBarWindow32Hwnd = FindWindowEx(ByVal FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "MSTaskSwWClass", "Running Applications"), _
                                    ByVal 0&, "ToolbarWindow32", "Running Applications")
    
        g_RunningProgramsHwnd = FindWindowEx(FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "MsTaskSwWClass", vbNullString), ByVal 0&, "ToolbarWindow32", vbNullString)
        If g_RunningProgramsHwnd = 0 Then
            g_RunningProgramsHwnd = FindWindowEx(FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "MSTaskSwWClass", vbNullString), ByVal 0&, "MSTaskListWClass", vbNullString)
        
            If g_RunningProgramsHwnd = 0 Then
                'Reset update trigger (forcing routine to later update again)
                g_TaskBarHwnd = -1
            End If
        End If
        
        g_StartButtonHwnd = FindWindowEx(g_TaskBarHwnd, 0, "Button", vbNullString)
        If g_StartButtonHwnd = 0 Then
            'Windows Vista/Seven
            
            g_StartButtonHwnd = FindWindow("Button", "Start")
            If g_StartButtonHwnd = 0 Then
                'Reset update trigger (forcing routine to later update again)
                g_TaskBarHwnd = -1
            Else
                g_WindowsVista = True
            End If
            
        End If
    
        updatedHwnd = True
    End If
    
    UpdatehWnds = updatedHwnd
End Function

Function IsTaskBarBehindWindow(hWnd As Long)
    
    If GetZOrder(g_TaskBarHwnd) > GetZOrder(hWnd) Then
        IsTaskBarBehindWindow = True
    Else
        IsTaskBarBehindWindow = False
    End If
    
End Function

Function IsWindowTopMost(hWnd As Long)

Dim windowStyle As Long

    IsWindowTopMost = False
    windowStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If IsStyle(windowStyle, WS_EX_TOPMOST) Then
        IsWindowTopMost = True
    End If

End Function

Public Function EnumWindowsAsTaskList(ByRef srcCollection As TaskList)
    On Error GoTo Handler

' Clear list, then fill it with the running
' tasks. Return the number of tasks.
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'

    If Not srcCollection Is Nothing Then
    
        Set m_TargetTaskList = srcCollection
        Call EnumWindows(AddressOf fEnumWindowsCallBack, ByVal 0)
    End If
    
    Exit Function
Handler:
    LogError Err.Number, "EnumerateWindowsAsTaskObject(); " & Err.Description, "TaskbarHelper"
    
End Function

Public Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long

    If IsVisibleToTaskBar(hWnd) Then
        m_TargetTaskList.AddWindowByHwnd hWnd
    End If

fEnumWindowsCallBack = True
End Function

Public Function IsVisibleToTaskBar(hWnd As Long) As Boolean

Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean

    IsVisibleToTaskBar = False
    
    ' This callback function is called by Windows (from
    ' the EnumWindows API call) for EVERY window that exists.
    ' It populates the listbox with a list of windows that we
    ' are interested in.
    '
    ' Windows to display are those that:
    '   -   are not this app's
    '   -   are visible
    '   -   do not have a parent
    '   -   have no owner and are not Tool windows OR
    '       have an owner and are App windows
    '       can be activated

    If IsWindowVisible(hWnd) Then
        If (GetParent(hWnd) = 0) Then
            bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
            
                IsVisibleToTaskBar = True
            End If
        End If
    End If
End Function

Public Function IsVisibleOnTaskbar(lExStyle As Long) As Boolean
    IsVisibleOnTaskbar = False

    If (lExStyle And WS_EX_APPWINDOW) Then
        IsVisibleOnTaskbar = True
    End If
End Function



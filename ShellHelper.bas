Attribute VB_Name = "ShellHelper"
Private Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop _
    As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Private shellWindowHandle As Long

Public Function HandleError(ByVal errorCode As Long)
    MsgBox "error retrieving handle: " & errorCode
End Function

Public Function FindShellWindow() As Long

'IntPtr progmanHandle;
'IntPtr defaultViewHandle = IntPtr.Zero;
'IntPtr workerWHandle;
'int errorCode = NativeMe

Dim progmanHandle As Long
Dim defaultViewHandle As Long
Dim workerWHandle As Long
Dim errorCode As Long
Dim ret As Long

    'Try the easy way first. "SHELLDLL_DefView" is a child window of "Progman".
    progmanHandle = FindWindowEx(0, 0, "Progman", vbNullString)
    If Not (progmanHandle = 0) Then
        defaultViewHandle = FindWindowEx(progmanHandle, 0, "SHELLDLL_DefView", vbNullString)
        errorCode = GetLastError()
    End If
    
    If Not (defaultViewHandle = 0) Then
        FindShellWindow = defaultViewHandle
        Exit Function
    ElseIf Not errorCode = ERROR_SUCCESS Then
        HandleError errorCode
        Exit Function
    End If
    
    'Try the not so easy way then. In some systems "SHELLDLL_DefView" is a child of "WorkerW"
    errorCode = ERROR_SUCCESS
    workerWHandle = FindWindowEx(0, 0, "WorkerW", vbNullString)
    Debug.Print "FindShellWindow.workerWHandle: " & workerWHandle

    If Not (workerWHandle = 0) Then
        defaultViewHandle = FindWindowEx(workerWHandle, 0, "SHELLDLL_DefView", vbNullString)
        errorCode = GetLastError()
    End If
    
    If Not (defaultViewHandle = 0) Then
        FindShellWindow = defaultViewHandle
        Exit Function
    ElseIf Not (errorCode = ERROR_SUCCESS) Then
        HandleError errorCode
        Exit Function
    End If
    
    shellWindowHandle = 0
    
    'Try the hard way. In some systems "SHELLDLL_DefView" is a child or a child of "Progman".
    If EnumChildWindows(progmanHandle, AddressOf EnumWindowsProc, ret) = 0 Then
        errorCode = GetLastError()
        If Not (errorCode = ERROR_SUCCESS) Then
            HandleError errorCode
            Exit Function
        End If
    End If

    'Try the even more harder way. Just in case "SHELLDLL_DefView" is in another desktop.
    If (shellWindowHandle = 0) Then
        If EnumDesktopWindows(0, AddressOf EnumWindowsProc, progmanHandle) Then
            errorCode = GetLastError()
            If Not (errorCode = ERROR_SUCCESS) Then
                HandleError errorCode
                Exit Function
            End If
        End If
    End If
    
    FindShellWindow = shellWindowHandle
End Function

Public Function EnumWindowsProc(ByVal handle As Long, ByVal lParam As Long) _
       As Long

    Dim foundHandle As Long
    foundHandle = FindWindowEx(handle, 0, "SHELLDLL_DefView", vbNullString)

    If Not (foundHandle = 0) Then
        shellWindowHandle = foundHandle
        
        EnumWindowsProc = 0
        Exit Function
    End If
    
    'Continue Enumerating
    EnumWindowsProc = 1
End Function




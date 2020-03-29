Attribute VB_Name = "MiscSupport"
Option Explicit

'Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function ChangeWindowMessageFilter Lib "user32.dll" (ByVal Message As Long, ByVal dwFlag As Integer) As Boolean
Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Long) As Long

Public Const MSGFLT_ADD = 1
Public Const MSGFLT_REMOVE = 2

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long

Public Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByRef oldValue As Long) As Long
Public Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByRef oldValue As Long) As Long

Private Declare Function MakeSureDirectoryPathExists Lib _
        "IMAGEHLP.DLL" (ByVal DirPath As String) As Long

Private Declare Function FindFirstFileW Lib "kernel32" _
 _
   (ByVal lpFileName As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Enum MNUITEMTYPE
    CHECKEDITEM = MF_CHECKED Or MF_STRING
    UNCHECKEDITEM = MF_UNCHECKED Or MF_STRING
    NORMALITEM = MF_STRING
End Enum

Public Function RefreshRelevantPads(ByVal targetPad As ViTab)

Dim thisWindow As ViPickWindow

    For Each thisWindow In ViPickWindows
        If thisWindow.UniqueID = targetPad.SharedViPadIdentifer Then
            thisWindow.RefreshContents
        End If
    Next

End Function

Public Function CloneViItem(ByRef sourceViPadItem As LaunchPadItem) As LaunchPadItem

Dim thisItem As New LaunchPadItem
    Set CloneViItem = thisItem

    With thisItem
        .Caption = sourceViPadItem.Caption
        Set .Icon = sourceViPadItem.Icon
        .IconPath = sourceViPadItem.IconPath
        .TargetPath = sourceViPadItem.TargetPath
        .TargetArguements = sourceViPadItem.TargetArguements
    End With

End Function


Public Function UnTopMostViPadWindows()

Dim thisForm As ViPickWindow
    For Each thisForm In ViPickWindows
        Call SetWindowPos(thisForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    Next

End Function

Public Function TopMostViPadWindows()

Dim thisForm As ViPickWindow
    For Each thisForm In ViPickWindows
        Call SetWindowPos(thisForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    Next

End Function

Public Function IsHwndBelongToUs(ByVal hWnd As Long) As Boolean

Dim thisForm As Form
    For Each thisForm In Forms
        If thisForm.hWnd = hWnd Then
            IsHwndBelongToUs = True
            Exit For
        End If
    Next

End Function

Function SetOwner(ByVal HwndtoUse, ByVal HwndofOwner) As Long
    SetOwner = SetWindowLong(HwndtoUse, GWL_HWNDPARENT, HwndofOwner)
End Function

Function ExistInCol(ByRef cTarget As Collection, sKey) As Boolean
    On Error GoTo Handler
    ExistInCol = Not (IsEmpty(cTarget(sKey)))
    
    Exit Function
Handler:
    ExistInCol = False
End Function

Function IsLeftMouseButtonDown() As Boolean
    IsLeftMouseButtonDown = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Public Function EnsureDirectoryExists(ByVal newPath) As Boolean
    Dim sPath As String
    'Add a trailing slash if none
    sPath = newPath & IIf(Right$(newPath, 1) = "\", "", "\")

    'Call API
    If MakeSureDirectoryPathExists(sPath) <> 0 Then
        'No errors, return True
        EnsureDirectoryExists = True
    End If

End Function

Public Function GetPngCodecCLSID() As CLSID

Dim thisCLSID As New GDIPImageEncoderList

    GetPngCodecCLSID = thisCLSID.EncoderForMimeType("image/png").CodecCLSID

End Function

Public Function ApplicationDataPath() As String
    ApplicationDataPath = Environ("appdata") & "\" & App.ProductName
End Function

Public Function ApplicationLnkBankPath() As String
    ApplicationLnkBankPath = ApplicationDataPath & "\lnks\"
End Function

Public Function Wow64Wrapper(ByVal sPath As String)

Dim theLnkTarget As String

    Wow64Wrapper = sPath

    If Is64bit Then
        theLnkTarget = ResolveLink(sPath)
    
        If theLnkTarget <> vbNullString And Not FileExists(theLnkTarget) Then
            If FileExists(Replace(theLnkTarget, Environ("ProgramFiles"), Environ("ProgramW6432"))) Then
                Wow64Wrapper = Replace(theLnkTarget, Environ("ProgramFiles"), Environ("ProgramW6432"))
            ElseIf FileExists(Replace(LCase(theLnkTarget), "system32", "sysnative")) Then
                Wow64Wrapper = Replace(LCase(theLnkTarget), "system32", "sysnative")
            End If
        End If
    End If

End Function

Public Function FileExists(sSource As String) As Boolean
    FileExists = FSO.FileExists(sSource)
End Function

Public Function ResolveLink(szPath As String) As String
    On Error GoTo Handler

Dim thisLnk As New ShellLinkClass
    thisLnk.Resolve szPath
    
    ResolveLink = thisLnk.Target
    
    Exit Function
Handler:
End Function

Public Function Is64bit() As Boolean
    Dim Handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    Handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If Handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim(StrIn)
   End Select
End Function

Public Function PrintHeader(ByVal theBank As ViBank)

Dim tabIndex As Long
Dim thisTab As ViTab
Dim printInfo As String

    For tabIndex = 1 To theBank.GetCollectionCount
        Set thisTab = theBank.GetTabByIndex(tabIndex)
        printInfo = printInfo & "[" & tabIndex & "]" & thisTab.Alias & " "
    Next

    Debug.Print printInfo
End Function

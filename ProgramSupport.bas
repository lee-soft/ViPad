Attribute VB_Name = "ProgramSupport"
Option Explicit

Public Declare Function TRACKMOUSEEVENT Lib "comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, _
ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public VIPAD_SPECIAL_URL As String

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100

Private Const CONTEXT_MENU As String = "Add to ViPad"

Private m_Registry As RegistryClass

Private m_launchPadIndex As Long
Private m_vipadInstanceIndex As Long
Private m_vipadKeyIndex As Long

Private m_wScriptShellObject As Object
Private m_defaultBrowserIcon As AlphaIcon

Public hwndDesktopChild As Long

Public Enum TrackMouseEventFlags
    TME_HOVER = 1&
    TME_LEAVE = 2&
    TME_NONCLIENT = &H10&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Const Toggle_Desktop_Icons As Long = &H7402

Public Function HTMLAsciiToChar(ByVal szData As String) As String
    On Error Resume Next

Dim thisChar As String
Dim charIndex As Long
Dim nextSemiColon As Long
Dim realCharacter As String

    Do
        charIndex = charIndex + 1
        thisChar = Mid(szData, charIndex, 1)
        
        If (thisChar = "&") Then
            nextSemiColon = InStr(charIndex, szData, ";")
            If nextSemiColon <= Len(szData) Then
                realCharacter = Chr(CLng(Mid(szData, charIndex + 2, (nextSemiColon - charIndex) - 2)))
                szData = Mid(szData, 1, charIndex - 1) & realCharacter & Mid(szData, nextSemiColon + 1)
            End If
        End If
        
    Loop While (charIndex < Len(szData))
    HTMLAsciiToChar = szData
End Function

Public Function GetWebpageTitle(ByRef szHtmlData As String) As String
    On Error Resume Next

Dim titleStart As Long
Dim titleEnd As String

    titleStart = InStr(szHtmlData, "<title>")
    If titleStart = 0 Then
        Exit Function
    End If
    titleStart = titleStart + Len("<title>")
    
    titleEnd = InStr(titleStart, szHtmlData, "</title>")
    If titleEnd > titleStart Then
        GetWebpageTitle = HTMLAsciiToChar(Mid(szHtmlData, titleStart, titleEnd - titleStart))
    End If
End Function

Public Function InIDE() As Boolean
'Returns whether we are running in vb(true), or compiled (false)
    Static counter As Variant
    If IsEmpty(counter) Then
        counter = 1
        Debug.Assert InIDE() Or True
        counter = counter - 1
    ElseIf counter = 1 Then
        counter = 0
    End If
    InIDE = counter
 
End Function

Public Function GetGlobalWScriptShellObject() As Object
    If m_wScriptShellObject Is Nothing Then
        Set m_wScriptShellObject = CreateObject("WScript.Shell")
    End If
    
    Set GetGlobalWScriptShellObject = m_wScriptShellObject
End Function

Public Function TrackMouse(hWnd As Long) As Boolean

Dim ET As TRACKMOUSEEVENT

    TrackMouse = False
    
    'initialize structure
    ET.cbSize = Len(ET)
    ET.hwndTrack = hWnd
    ET.dwFlags = TME_LEAVE
    'start the tracking
    If Not TRACKMOUSEEVENT(ET) = 0 Then
        TrackMouse = True
    End If
    
End Function

Public Function UnInstallApp()

    On Error GoTo Handler
    
    UnInstallRegistryIfNeeded

Dim Shell, DesktopPath, link
Dim theLnkFile As String

    Set Shell = CreateObject("WScript.Shell")
    DesktopPath = CStr(Shell.SpecialFolders("Desktop"))
    theLnkFile = DesktopPath & "\" & App.ProductName & ".lnk"
    
    If FileExists(theLnkFile) = False Then
        Set link = Shell.CreateShortcut(theLnkFile)
    
        link.TargetPath = App.Path & "\" & App.EXEName & ".exe"
        link.WorkingDirectory = App.Path
        
        link.Save
    End If
    
    If WindowsVersion.dwMajorVersion = 6 And WindowsVersion.dwMinorVersion >= 1 Then
        UnInstallLnkToTaskBar theLnkFile
    End If
    
    Kill theLnkFile
    Exit Function
Handler:
End Function

Public Function InstallApp()
    On Error GoTo Handler

Dim Shell, DesktopPath, link
Dim theLnkFile As String

    Set Shell = CreateObject("WScript.Shell")
    DesktopPath = CStr(Shell.SpecialFolders("Desktop"))
    theLnkFile = DesktopPath & "\" & App.ProductName & ".lnk"
    
    If FileExists(theLnkFile) = False Then
        Set link = Shell.CreateShortcut(theLnkFile)
    
        link.TargetPath = App.Path & "\" & App.EXEName & ".exe"
        link.WorkingDirectory = App.Path
        
        link.Save
    End If
    
    If WindowsVersion.dwMajorVersion = 6 And WindowsVersion.dwMinorVersion >= 1 Then
        InstallLnkToTaskBar theLnkFile
    End If
    
    Exit Function
Handler:
End Function

Private Function ApplySuperBarOperation(theLnkFile As String, theOperation As String)

    On Error GoTo Handler

Dim ShellApp, ParentFolder
Set ShellApp = CreateObject("Shell.Application")
'Set FSO = CreateObject("Scripting.FileSystemObject")

Dim LnkFile
Dim LnkFileName As String

    Set LnkFile = FSO.GetFile(theLnkFile)
    Set ParentFolder = ShellApp.NameSpace(Replace(LnkFile.Path, "\" & LnkFile.Name, ""))
    
    LnkFileName = Replace(LCase(FSO.GetFile(LnkFile).Name), ".lnk", "")

If (FSO.FileExists(theLnkFile)) Then
    Dim tmp, verb
    Dim desktopImtes, item
    Set desktopImtes = ParentFolder.Items()

    For Each item In desktopImtes
        If (LCase(item.Name) = LnkFileName) Then
            For Each verb In item.Verbs
                If (verb.Name = theOperation) _
        Then 'If (verb.Name = "??????(&K)")
                    verb.DoIt
                End If
            Next
        End If
    Next

End If

Set ShellApp = Nothing

    Exit Function
Handler:
End Function

Private Function UnInstallLnkToTaskBar(theLnkFile As String)
    ApplySuperBarOperation theLnkFile, "Unpin from Tas&kbar"
End Function

Private Function InstallLnkToTaskBar(theLnkFile As String)
    ApplySuperBarOperation theLnkFile, "Pin to Tas&kbar"
End Function

Public Function GetWindowTitle(ByVal targethWnd As Long) As String

' Display the text of the title bar of window Form1
Dim textlen As Long ' receives length of text of title bar
Dim titlebar As String ' receives the text of the title bar
Dim slength As Long ' receives the length of the returned string

    ' Find out how many characters are in the window's title bar
    textlen = GetWindowTextLength(targethWnd)
    titlebar = Space(textlen + 1) ' make room in the buffer, allowing for the terminating null character
    slength = GetWindowText(targethWnd, titlebar, textlen + 1) ' read the text of the window
    GetWindowTitle = Left(titlebar, slength) ' extract information from the buffer

End Function

Public Function IsGlassAvailable() As Boolean

    If WindowsVersion.dwMajorVersion > 5 Then
        IsGlassAvailable = True
    End If

End Function

Public Function UnApplyGlassIfPossible(hWnd As Long) As Boolean

    If IsGlassAvailable Then
        UnApplyGlassIfPossible = True

        Dim theRect As RECTL
        Win32.ApplyGlass hWnd, theRect
    End If
    
End Function

Public Function ApplyGlassIfPossible(hWnd As Long) As Boolean

    If IsGlassAvailable Then
        Dim negativeRect As RECTL
        negativeRect.Left = -1
        negativeRect.Top = -1
            
        If Win32.ApplyGlass(hWnd, negativeRect) = S_OK Then
            ApplyGlassIfPossible = True
        End If
    End If
    
End Function

Public Function GetWindowsOSVersion() As OSVERSIONINFO

Dim osv As OSVERSIONINFO
    osv.dwOSVersionInfoSize = Len(osv)
    
    If GetVersionEx(osv) = 1 Then
        GetWindowsOSVersion = osv
    End If

End Function

Public Function GetPublicDesktopPath() As String

    On Error GoTo Handler

Dim WshShell As Object

Set WshShell = CreateObject("WScript.Shell")
GetPublicDesktopPath = WshShell.SpecialFolders("AllUsersDesktop")

    Exit Function
Handler:
    GetPublicDesktopPath = Environ("public") & "\desktop\"


End Function

Public Function GetUserDesktopPath() As String

    On Error GoTo Handler

Dim WshShell As Object

Set WshShell = CreateObject("WScript.Shell")
GetUserDesktopPath = WshShell.SpecialFolders("Desktop")

    Exit Function
Handler:
    GetUserDesktopPath = Environ("userprofile") & "\desktop\"

End Function

Public Function DefaultBrowserIcon() As AlphaIcon
    
    If Not m_defaultBrowserIcon Is Nothing Then
        Set DefaultBrowserIcon = m_defaultBrowserIcon
        Exit Function
    End If
    
Set m_defaultBrowserIcon = New AlphaIcon
    Set DefaultBrowserIcon = m_defaultBrowserIcon

Dim theIcon As Long
Dim imageFileNamePath As String

    theIcon = IconHelper.GetIconFromFile(DefaultBrowserIconPath, SHIL_JUMBO)
    If IconHelper.IconIs48(DefaultBrowserIconPath) Then
        theIcon = IconHelper.GetIconFromFile(DefaultBrowserIconPath, SHIL_EXTRALARGE)
    End If
    
    m_defaultBrowserIcon.CreateFromHICON theIcon
End Function

Public Function DefaultBrowserIconPath() As String

Dim returnString As String
Dim commaLocation As Long

    InitializeRegistryIfUnitialized

    returnString = m_Registry.GetStringValue(HKEY_CURRENT_USER, "SOFTWARE\Classes\http\DefaultIcon", "")
    
    If Len(returnString) > 0 And InStr(returnString, ".exe") > 0 Then
        commaLocation = InStr(returnString, ",")
        
        If commaLocation > 0 Then
            DefaultBrowserIconPath = Mid(returnString, 1, commaLocation - 1)
        End If
    Else
        returnString = Trim(DefaultBrowserCommand())
        If Len(returnString) > 0 Then
            If Left(returnString, 1) = """" Then
                commaLocation = InStr(2, returnString, """")

                If commaLocation > 0 Then
                    DefaultBrowserIconPath = Mid(returnString, 2, commaLocation - 2)
                End If
            Else
                DefaultBrowserIconPath = returnString
            End If
        End If
    End If
End Function

Public Function DefaultBrowserCommand() As String
    InitializeRegistryIfUnitialized
    
    DefaultBrowserCommand = m_Registry.GetStringValue(HKEY_CURRENT_USER, "SOFTWARE\Classes\http\shell\open\command", "")
End Function

Private Function UnhandleType(theType As String, Optional theText As String = CONTEXT_MENU)
    InitializeRegistryIfUnitialized
    
Dim strKeyValue As String
    strKeyValue = App.Path & "\" & App.EXEName & ".exe" & " " & """" & "%1" & """"

    If m_Registry.KeyExists(HKEY_CLASSES_ROOT, theType & "\shell\" & theText) Then
        m_Registry.DeleteKey HKEY_CLASSES_ROOT, theType & "\shell\" & theText
    End If

End Function

Private Function HandleType(theType As String, Optional theText As String = CONTEXT_MENU)
    InitializeRegistryIfUnitialized
    
Dim strKeyValue As String
    strKeyValue = App.Path & "\" & App.EXEName & ".exe" & " " & """" & "%1" & """"

    If Not m_Registry.KeyExists(HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command") Then
        m_Registry.CreateKey HKEY_CLASSES_ROOT, theType & "\shell\" & theText
        m_Registry.CreateKey HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command"
        
        m_Registry.SetStringValue HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command", "", strKeyValue
    Else
        If m_Registry.GetStringValue(HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command\", "") <> strKeyValue Then
            m_Registry.SetStringValue HKEY_CLASSES_ROOT, theType & "\shell\" & theText & "\command", "", strKeyValue
        End If
    End If

End Function

Public Function UnInstallRegistryIfNeeded()

    InitializeRegistryIfUnitialized

    UnhandleType "lnkfile"
    UnhandleType "*"
    
End Function

Public Function InstallRegistryIfNeeded()

    InitializeRegistryIfUnitialized

    HandleType "lnkfile"
    HandleType "*"

End Function

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
   MAKELPARAM = MakeLong(wLow, wHigh)
End Function

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    If fOnTop Then
        Call SetWindowPos(frmForm.hWnd, HWND_TOPMOST, iLeft, iTop, iWidth, iHeight, 0)
    Else
        Call SetWindowPos(frmForm.hWnd, HWND_NOTOPMOST, iLeft, iTop, iWidth, iHeight, 0)
    End If
    
End Sub

' Set or clear a style or extended style value.
Public Function SetWindowStyle(ByVal hWnd As Long, ByVal _
    extended_style As Boolean, ByVal style_value As Long, _
    ByVal new_value As Boolean, ByVal bRefresh As Boolean)
    
Dim style_type As Long
Dim style As Long
   
    If extended_style Then
        style_type = GWL_EXSTYLE
    Else
        style_type = GWL_STYLE
    End If

    ' Get the current style.
    style = GetWindowLong(hWnd, style_type)

    ' Add or remove the indicated value.
    If new_value Then
        style = style Or style_value
    Else
        style = style And Not style_value
    End If
    
    ' Hide Window if Changing ShowInTaskBar
    If bRefresh Then
        ShowWindow hWnd, SW_HIDE
    End If

    ' Set the style.
    SetWindowLong hWnd, style_type, style

    ' Show Window if Changing ShowInTaskBar
    If bRefresh Then
        ShowWindow hWnd, SW_SHOW
    End If
    
    ' Make the window redraw.
    SetWindowPos hWnd, 0, 0, 0, 0, 0, _
        SWP_FRAMECHANGED Or _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOREPOSITION Or _
        SWP_NOZORDER
End Function

Public Function GetFileLocation(thePath As String) As String
    InitializeFSOIfUnitialized

Dim thisFile As Scripting.File
    
    If Not FSO.FileExists(thePath) Then
        Exit Function
    End If
    
    Set thisFile = FSO.GetFile(thePath)
    GetFileLocation = thisFile.ParentFolder.Path
    
End Function

Public Function GetDefViewWnd() As Long
Dim hDefViewWnd As Long

    EnumWindows AddressOf FindDLV, hDefViewWnd
    GetDefViewWnd = hDefViewWnd
End Function

Private Function FindDLV(ByVal hWndPM As Long, ByRef lParam As Long) As Long

Dim hWnd As Long
    FindDLV = APITRUE

    hWnd = FindWindowEx(hWndPM, 0, "SHELLDLL_DefView", vbNullString)
    If hWnd <> 0 Then
        FindDLV = APIFALSE
        lParam = hWnd
    End If

End Function

Public Function IsDesktopIconsVisible() As Boolean
    InitializeRegistryIfUnitialized
    
    If m_Registry.GetLongValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideIcons") = 0 Then
        IsDesktopIconsVisible = True
    End If
End Function

Private Function GetDesktopSystemListHwnd() As Long

Dim hDefViewWnd As Long

    EnumWindows AddressOf FindDLV, hDefViewWnd
    GetDesktopSystemListHwnd = FindWindowEx(hDefViewWnd, 0, "SysListView32", vbNullString)
End Function

Public Function HideDesktopIcons_Windows8()
    ShowWindow GetDesktopSystemListHwnd(), SW_HIDE
End Function

Public Function ShowDesktopIcons_Windows8()
    ShowWindow GetDesktopSystemListHwnd(), SW_SHOW
End Function

Public Function ToggleDesktopIcons() As Boolean

    PostMessage ByVal hwndDesktopChild, ByVal &H111, ByVal Toggle_Desktop_Icons, ByVal 0
End Function

Public Function HideDesktopIcons()
    If WindowsVersion.dwMajorVersion = 6 And WindowsVersion.dwMinorVersion = 2 Then
        HideDesktopIcons_Windows8
        Exit Function
    End If
    
    If IsDesktopIconsVisible Then
        ToggleDesktopIcons
    End If
End Function

Public Function ShowDesktopIcons()
    If WindowsVersion.dwMajorVersion = 6 And WindowsVersion.dwMinorVersion = 2 Then
        ShowDesktopIcons_Windows8
        Exit Function
    End If

    If Not IsDesktopIconsVisible Then
        ToggleDesktopIcons
    End If
End Function

Public Function GetNextViPadKey() As String
    m_vipadKeyIndex = m_vipadKeyIndex + 1
    GetNextViPadKey = "vipadkey_" & m_vipadKeyIndex
End Function

Public Function GetNextViPadInstanceID() As String
    m_vipadInstanceIndex = m_vipadInstanceIndex + 1
    GetNextViPadInstanceID = "vipad_" & m_vipadInstanceIndex
End Function

Public Function GetNextLaunchPadID() As String
    m_launchPadIndex = m_launchPadIndex + 1
    GetNextLaunchPadID = "obj_" & m_launchPadIndex
End Function

Public Function QuoteString(ByVal szSourceString) As String
    QuoteString = """" & szSourceString & """"
End Function

Public Function GetVersionNumber() As String
    GetVersionNumber = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function CopyLnkToBank(ByVal szSourceLinkPath As String) As String
    On Error GoTo Handler

Dim szNewLnkPath As String

    InitializeFSOIfUnitialized
    
    szNewLnkPath = GenerateAvailableLnkFileName
    If szNewLnkPath = vbNullString Then Exit Function
    
    If CopyFile(szSourceLinkPath, szNewLnkPath, APITRUE) = 0 Then
        Exit Function
    End If
    
    CopyLnkToBank = szNewLnkPath

Handler:
End Function

Public Function GenerateAvailableLnkFileName()
    On Error GoTo Handler

Dim theFile As Scripting.File
Dim proposedFileName As String
Dim iconIndex As Long
    
    InitializeFSOIfUnitialized
    
    Do
        iconIndex = iconIndex + 1
        proposedFileName = ApplicationLnkBankPath & "lnk_" & iconIndex & ".lnk"
    Loop Until FSO.FileExists(proposedFileName) = False
    
    GenerateAvailableLnkFileName = proposedFileName

    Exit Function
Handler:
End Function

Public Function GenerateAvailableFileName()
    On Error GoTo Handler

Dim theFile As Scripting.File
Dim proposedFileName As String
Dim iconIndex As Long
    
    InitializeFSOIfUnitialized
    
    Do
        iconIndex = iconIndex + 1
        proposedFileName = ApplicationDataPath & "\icon_" & iconIndex & ".png"
    Loop Until FSO.FileExists(proposedFileName) = False
    
    GenerateAvailableFileName = proposedFileName

    Exit Function
Handler:
End Function

Public Function GetFileNameFromPath(ByVal szPath As String)
    On Error GoTo Handler
    
    InitializeFSOIfUnitialized
    
    If FSO.FileExists(szPath) Then
        GetFileNameFromPath = FSO.GetFile(szPath).Name
    ElseIf FSO.FolderExists(szPath) Then
        GetFileNameFromPath = FSO.GetFolder(szPath).Name
    End If

    Exit Function
Handler:
End Function

Public Function FWriteBinary(szPath As String, szData As String)

Dim fnum As Long

    On Error Resume Next
    Kill szPath
    
    ' Save the file.
    fnum = FreeFile
    Open szPath For Binary As #fnum
        Put #fnum, 1, szData
    Close fnum

End Function

Public Function FWrite(szPath As String, szData As String) As Boolean
    'On Error GoTo Handler
    
Dim theStream As Scripting.TextStream
    
    InitializeFSOIfUnitialized
    
    If FileExists(szPath) Then
        FSO.DeleteFile szPath, True
    End If
    
    Set theStream = FSO.OpenTextFile(szPath, ForWriting, True)
    theStream.Write szData
    
    FWrite = True
    Exit Function
Handler:
    FWrite = False
End Function

Public Function Fload(szPath As String)
    On Error GoTo Handler

Dim theFile As Scripting.File
Dim theStream As Scripting.TextStream
    
    InitializeFSOIfUnitialized
    
    Set theFile = FSO.GetFile(szPath)
    Set theStream = theFile.OpenAsTextStream(ForReading)
    
    Fload = theStream.ReadAll()
    Exit Function
Handler:
End Function

Private Function InitializeFSOIfUnitialized()
    If FSO Is Nothing Then
        Set FSO = New FileSystemObject
    End If
End Function

Private Function InitializeRegistryIfUnitialized()
    If m_Registry Is Nothing Then
        Set m_Registry = New RegistryClass
    End If
End Function

Public Function GetProgramExecutables()
    
Dim programFolder As Folder
Dim programs As New Collection
    
    InitializeFSOIfUnitialized
    
    'Set programFolder = FSO.GetFolder(GetShellVariable("programs"))
    'PopulateCollection programs, programFolder
    Set programFolder = FSO.GetFolder(GetShellVariable("Common Programs"))
    PopulateCollection programs, programFolder

    Set GetProgramExecutables = programs
End Function

Function GetShellVariable(szVarName As String) As String
    InitializeRegistryIfUnitialized
    
    GetShellVariable = m_Registry.GetStringValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", szVarName, "")
    If Not GetShellVariable = "" Then Exit Function
    
    GetShellVariable = m_Registry.GetStringValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", szVarName, "")
    
End Function

Private Function PopulateCollection(ByRef sourceCollection As Collection, ByRef sourceFolder As Scripting.Folder)

Dim thisFile As Scripting.File
Dim thisFolder As Scripting.Folder
Dim thisLnk As ShellLinkClass
Dim thisProgram As Program

    If sourceCollection Is Nothing Then Exit Function
    
    For Each thisFile In sourceFolder.Files
        If Not thisFile.Attributes And Hidden Then
            Set thisLnk = New ShellLinkClass
            thisLnk.Resolve thisFile.Path

            If LCase(Right(thisLnk.Target, 4)) = ".exe" Then
                Set thisProgram = New Program
                thisProgram.PhysicalTarget = thisLnk.Target
                thisProgram.FileName = Replace(thisFile.Name, ".lnk", "")
                thisProgram.IconLocation = thisFile.Path

                sourceCollection.Add thisProgram
            End If
            'cRes.createNode ExtOrNot(thisFile.Name), MakeSearchable(thisFile.Name), thisFile.Path, GetIconY(thisFile.Path), Replace(thisFile.Path, stripperKey, ""), True
        End If
    Next
    
    For Each thisFolder In sourceFolder.SubFolders
        If Not thisFolder.Attributes And Hidden Then
            PopulateCollection sourceCollection, thisFolder
        End If
    Next

End Function

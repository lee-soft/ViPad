Attribute VB_Name = "MainMod"
Option Explicit

'==========================================APIs==========================================
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private m_GDIInitialized As Boolean

Public g_hiddenDesktop As Boolean
Public g_mainHwnd As Long
Public Const g_APPID As String = "VIPAD27081987"

Public Config As ViSettings
Public AddDesktopIcons As Boolean
Public WindowsVersion As OSVERSIONINFO

Public PublicItemBank As ViBank
Public ViPickWindows As Collection
Public SystemNotifications As SystemNotificationManager
Public AppTrayIcon As TrayIcon
Public AppZorderKeeper As ZOrderContainer
Public AppOptionsWindow As OptionsWindow
Public URLCatcher As URLNotification
Public FSO As FileSystemObject

Private m_CmdLine As CommandLine
Private m_ending As Boolean
Private m_addLinkFileWhenLoaded As String

Public AppRestarting As Boolean
Public BlockBottomMost As Boolean

Public Type argb
    R As Byte
    G As Byte
    b As Byte
    A As Byte
End Type

Private Const EXIT_PROGRAM As Long = 1

Public Function StripExtension(ByVal szFileName As String) As String

Dim szExtension As String
    
    If InStr(szFileName, ".") = 0 Then
        StripExtension = szFileName
        Exit Function
    End If
    
    szExtension = Mid(szFileName, InStrRev(szFileName, ".") + 1)
    
    If InStr(szExtension, " ") > 0 Then
        StripExtension = szFileName
    Else
        StripExtension = Mid(szFileName, 1, InStrRev(szFileName, ".") - 1)
    End If
End Function

Public Sub Long2ARGB(ByVal LongARGB As Long, ByRef argb As argb)
    CopyMemory argb, LongARGB, 4
End Sub

Public Function InitializeGDIIfNotInitialized() As Boolean
    
    If Not m_GDIInitialized Then
        ' Must call this before using any GDI+ call:
        If Not (gdipluswrapper.GDIPlusCreate(True)) Then
            Exit Function
        End If
    
        m_GDIInitialized = True
    End If
    
    InitializeGDIIfNotInitialized = m_GDIInitialized
End Function

Public Function PreviousInstanceHandler() As Boolean

Dim hwndPrevInstance As Long
Dim tCDS As COPYDATASTRUCT
Dim dataToSend() As Byte
    
    hwndPrevInstance = FindWindow("ThunderRT6FormDC", g_APPID)
    If hwndPrevInstance > 0 Then
        PreviousInstanceHandler = True
        dataToSend = GetFirstCommandIfAny

        With tCDS
            tCDS.lpData = VarPtr(dataToSend(0))
            tCDS.dwData = 88
            tCDS.cbData = UBound(dataToSend)
        End With
        
        SendMessage hwndPrevInstance, WM_COPYDATA, 0, tCDS
        Exit Function
    Else
        If App.PrevInstance Then
            Sleep 1000
            PreviousInstanceHandler = PreviousInstanceHandler
        End If
    End If

End Function

Private Function HandleSplash() As Long

Dim A As New XMLWindow
    
    If Not A.GoToURL("res://default_splash") Then
        Unload A
        Exit Function
    End If
    
    'If A.GoToURL("http://lee-soft.com/vipad/splash_screens/test.xml") = False Then
        'A.GoToURL "res://default_splash"
    'End If
    
    A.Show vbModal
    HandleSplash = A.returnCode

End Function

Private Function OleFileDrag(ByVal theFilePath As String)
    Dim theNewImage As New FloatingImage
    
    theNewImage.MoveToCursor
    theNewImage.Show
    'theNewImage.Height = 0
    theNewImage.FilePath = theFilePath
    theNewImage.OLEDrag
    
    Unload theNewImage
    Set theNewImage = Nothing
End Function

Private Function ShowDragImage(ByVal theSize As Long, theImagePath As String)
    Dim theNewImage As New FloatingImage
    Dim newImage As New GDIPImage
    newImage.FromFile theImagePath
    
    theNewImage.SetImage newImage
    theNewImage.Show
    theNewImage.TrackCursor
    
End Function

Private Function DetermineAction() As Long
    Set m_CmdLine = New CommandLine
    
    If m_CmdLine.Arguments > 0 Then
        Select Case UCase(m_CmdLine.Argument(1))
        
        Case "/UNINSTALL"
            UnInstallApp
            DetermineAction = EXIT_PROGRAM
        
        Case "/OLEDRAG"
            If m_CmdLine.Arguments > 1 Then
                OleFileDrag CStr(m_CmdLine.Argument(2))
                DetermineAction = EXIT_PROGRAM
            End If
        
        Case "/SHOWDRAGIMAGE"
            If m_CmdLine.Arguments > 2 Then
                ShowDragImage CLng(m_CmdLine.Argument(2)), CStr(m_CmdLine.Argument(3))
                DetermineAction = EXIT_PROGRAM
            End If
            
        Case Else
            m_addLinkFileWhenLoaded = m_CmdLine.Argument(1)
        
        End Select
    End If
End Function

Public Sub Main()

Dim ViPick As ViPickWindow

    If PreviousInstanceHandler Then
        Exit Sub
    End If
    
    Dim X As Long
    X = InitCommonControls
    InitializeGDIIfNotInitialized
    WindowsVersion = GetWindowsOSVersion
    
    If DetermineAction = EXIT_PROGRAM Then
        Exit Sub
    End If
    
    Set Config = New ViSettings
    InitializeSharedObjects

    EnsureDirectoryExists ApplicationDataPath
    EnsureDirectoryExists ApplicationLnkBankPath
    
    If Not InIDE Then
        If Config.SplashUpdater Then
            If HandleSplash() <> 0 Then
                Exit Sub
            End If
        End If

    End If
    
    PluginsHelper.Init
    
    If Not FileExists(MiscSupport.ApplicationDataPath & "\settings.xml") Then
        InstallApp
        
        FirstTimeWarning.Show vbModal
        AddDesktopIcons = FirstTimeWarning.AddDesktopShortcuts
    End If

    ActionSettings
    InstallRegistryIfNeeded
    
    If ViPickWindows.Count > 0 Then
        Set ViPick = ViPickWindows(1)
        If Not ViPick Is Nothing Then If AddDesktopIcons Then ViPick.AddShortcutsOnDesktop
    End If
    
    If FileExists(m_addLinkFileWhenLoaded) Then
        AppTrayIcon.Notify_AddNewLink m_addLinkFileWhenLoaded
    End If
End Sub

Private Function InitializeSharedObjects()
    'IMPORTANT DO BEFORE ZORDERKEEPER
    
    VIPAD_SPECIAL_URL = App.Path & "\ViPad_Special_Link.url"
    
    'hwndDesktopChild = GetWindow(FindWindow("Progman", "Program Manager"), GW_CHILD)
    hwndDesktopChild = FindShellWindow()

    Set FSO = New FileSystemObject
    Set PublicItemBank = ProgramIO.LoadBankXML(ApplicationDataPath & "\pads.xml")
    Set ViPickWindows = New Collection
    Set SystemNotifications = New SystemNotificationManager
    Set AppOptionsWindow = New OptionsWindow
    Set URLCatcher = New URLNotification
    Load URLCatcher
    
    Set AppTrayIcon = New TrayIcon
    Load AppTrayIcon

    Set AppZorderKeeper = New ZOrderContainer
    Load AppZorderKeeper
    
    If PublicItemBank.GetCollectionCount < 1 Then
        PublicItemBank.AddNewCollection "New Tab"
    End If
End Function

Private Function PositionMainWindow(ByRef ViPick As Form)
    If PublicItemBank Is Nothing Then Exit Function
    
    If PublicItemBank.GetCollectionCount > 0 Then
        Dim thisTab As ViTab
        Set thisTab = PublicItemBank.GetTabByIndex(1)
        
        Debug.Print "thisTab.Dimensions:: " & thisTab.Dimensions
        PositionWindow ViPick, Unserialize_RectL(thisTab.Dimensions)
    End If
End Function

Private Function OpenPadsForEachTab()
    
Dim headerItems As Collection

Dim thisCollectionIndex As Long
Dim thisWindowDimensions As GdiPlus.RECTL

Dim thisTab As ViTab
Dim tabIndex As Long

    Set headerItems = New Collection
    tabIndex = 1
    
    For thisCollectionIndex = 1 To PublicItemBank.GetCollectionCount

        Set thisTab = PublicItemBank.GetTabByIndex(thisCollectionIndex)

        Dim thisInstance As ViPickWindow
        Set thisInstance = New ViPickWindow
        Load thisInstance
  
        thisInstance.Tag = thisTab.Alias
        thisInstance.SetSelectedTab thisTab
        
        Debug.Print "thisTab.Dimensions:: " & thisTab.Dimensions
        thisWindowDimensions = Unserialize_RectL(thisTab.Dimensions)

        
        If Not IsRectL_Empty(thisWindowDimensions) Then
            PositionWindow thisInstance, thisWindowDimensions
            thisInstance.LayeredWindowHandler
            PositionWindow thisInstance, thisWindowDimensions
            
            
        Else
            thisInstance.Left = ((tabIndex * 30) * Screen.TwipsPerPixelX)
            thisInstance.Top = ((tabIndex * 30) * Screen.TwipsPerPixelY)
            
            tabIndex = tabIndex + 1
        End If
        
        thisInstance.Show
        
        'thisInstance.GlassContainer.UpdateClientPosition
        
        'AddItemHeader thisCollectionAlias, False
    Next
End Function

Private Function ForceVBCompleteFormInitialize(ByRef theForm As Form)
    theForm.Show
    theForm.Top = -200
    theForm.Left = -200
    theForm.Height = 200
    theForm.Width = 200
    theForm.Show
    theForm.ZOrder
End Function

Private Function ValidateDimensions(thisWindowDimensions As GdiPlus.RECTL)

Dim screenWidthPx As Long
Dim screenHeightPx As Long

    screenWidthPx = Screen.Width / Screen.TwipsPerPixelX
    screenHeightPx = Screen.Height / Screen.TwipsPerPixelY

    If (thisWindowDimensions.Left + thisWindowDimensions.Width) < 300 Or _
        (thisWindowDimensions.Top + thisWindowDimensions.Height) < 300 And _
        (thisWindowDimensions.Left + thisWindowDimensions.Width) > screenWidthPx Or _
        (thisWindowDimensions.Top + thisWindowDimensions.Height) > screenHeightPx Then

        thisWindowDimensions.Width = 800
        thisWindowDimensions.Height = 600
    
        thisWindowDimensions.Left = ((Screen.Width / Screen.Width / 2) / Screen.TwipsPerPixelX) + _
                                        (thisWindowDimensions.Width / 2)
        thisWindowDimensions.Top = ((Screen.Height / Screen.Height / 2) / Screen.TwipsPerPixelY) + _
                                        (thisWindowDimensions.Height / 2)
    End If


End Function

Private Function PositionWindow(ByRef theWindow As ViPickWindow, thisWindowDimensions As GdiPlus.RECTL)

    ValidateDimensions thisWindowDimensions

    theWindow.Width = thisWindowDimensions.Width * Screen.TwipsPerPixelX
    theWindow.Height = thisWindowDimensions.Height * Screen.TwipsPerPixelY
    theWindow.Left = thisWindowDimensions.Left * Screen.TwipsPerPixelX
    theWindow.Top = thisWindowDimensions.Top * Screen.TwipsPerPixelY

End Function

Private Function ActionSettings()
    If Config.HideDesktopOnBoot Then
        g_hiddenDesktop = True
        ProgramSupport.HideDesktopIcons
    End If
    
    If Config.InstanceMode Then
        OpenPadsForEachTab
    Else
        Dim ViPick As New ViPickWindow
        g_mainHwnd = ViPick.hWnd
        
        PositionMainWindow ViPick
        ViPick.LayeredWindowHandler
        PositionMainWindow ViPick
        
        
        ViPick.Show
        ViPick.ShowMultitabInterface
    End If
End Function

Public Function DumpPads()
    On Error GoTo Handler

Dim thisPadFile As DOMDocument

    Set thisPadFile = GetPadsXMLDocument(PublicItemBank)
    thisPadFile.Save MiscSupport.ApplicationDataPath & "\pads.xml"

    Exit Function
Handler:
    MsgBox Err.Description, vbCritical, "File I/O"
End Function

Private Sub UpdateRegisteredWindowDimensions()

Dim thisViPickWindow As ViPickWindow
    For Each thisViPickWindow In ViPickWindows
        thisViPickWindow.UpdateTabDimensionsIfPossible
    Next

End Sub

Public Function ToggleAppType()

Dim newModeTabbed As Boolean
Dim thisForm As ViPickWindow

    newModeTabbed = Not Config.InstanceMode

    AppRestarting = True

    UpdateRegisteredWindowDimensions
    DumpPads
    
    For Each thisForm In ViPickWindows
        Unload thisForm
    Next
    
    Config.InstanceMode = newModeTabbed
    
    ActionSettings
    
    AppRestarting = False
End Function

Public Function EndApplication()

Dim thisForm As ViPickWindow

    If m_ending Then Exit Function
    m_ending = True
    
    PluginsHelper.NotifyFriendsAppShutdown
    
    UpdateRegisteredWindowDimensions
    DumpPads

    For Each thisForm In ViPickWindows
        Unload thisForm
    Next
    
    If Not URLCatcher Is Nothing Then Unload URLCatcher
    
    'On Error Resume Next
    Set MainMod.SystemNotifications = Nothing
    
    If Not MainMod.AppZorderKeeper Is Nothing Then
        Unload MainMod.AppZorderKeeper
        Set MainMod.AppZorderKeeper = Nothing
    End If
    
    If Not MainMod.AppTrayIcon Is Nothing Then
        Unload MainMod.AppTrayIcon
        Set MainMod.AppTrayIcon = Nothing
    End If
    
    If Not MainMod.AppOptionsWindow Is Nothing Then
        Unload MainMod.AppOptionsWindow
        Set MainMod.AppOptionsWindow = Nothing
    End If
    
    If g_hiddenDesktop Then
        ProgramSupport.ShowDesktopIcons
    End If

    UnInstallRegistryIfNeeded
    
    Set Config = Nothing
End Function

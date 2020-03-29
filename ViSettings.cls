VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MinimizeAfterLauch As Boolean
Public HideDesktopOnBoot As Boolean
Public ShowControlBox As Boolean
Public TopMostWindow As Boolean
Public ShowOnTaskBar As Boolean
Public IconSize As Long
Public StickToDesktop As Boolean
Public DockMode As Boolean
Public SplashUpdater As Boolean
Public VisibleToTaskBar As Boolean
Public ForceLayeredMode As Boolean
Public InstanceMode As Boolean
Public ShowTrayIcon As Boolean
Public CatchWebLinks As Boolean
Public MiddleMouseActivation As Boolean

Public WindowState As FormWindowStateConstants

Private m_MainWindowDimensions As GdiPlus.RECTL

Public Property Get MainWindowRect() As GdiPlus.RECTL
    MainWindowRect = m_MainWindowDimensions
End Property

Public Property Let MainWindowRect(newRect As GdiPlus.RECTL)
    m_MainWindowDimensions = newRect
End Property

Private Function ReadSettings()
    On Error GoTo Handler

Dim settingsXml As New DOMDocument
Dim thisChild As IXMLDOMElement
    
    'Set Defaults
    IconSize = 70
    MinimizeAfterLauch = True
    DockMode = False
    ShowTrayIcon = True
    SplashUpdater = True
    VisibleToTaskBar = True
    WindowState = vbNormal
    CatchWebLinks = True
    MiddleMouseActivation = True
    
    If settingsXml.Load(MiscSupport.ApplicationDataPath & "\settings.xml") = False Then
        Exit Function
    End If
    
    For Each thisChild In settingsXml.selectSingleNode("settings").childNodes
        Select Case thisChild.tagName
        
        Case "MinimizeAfterLauch"
            MinimizeAfterLauch = CBool(thisChild.Text)
        
        Case "HideDesktopOnBoot"
            HideDesktopOnBoot = CBool(thisChild.Text)
        
        Case "ShowControlBox"
            ShowControlBox = CBool(thisChild.Text)
            
        Case "TopMostWindow"
            TopMostWindow = CBool(thisChild.Text)
            
        Case "StickToDesktop"
            StickToDesktop = CBool(thisChild.Text)
            
        Case "IconSize"
            IconSize = CLng(thisChild.Text)
            
        Case "DockMode"
            DockMode = CBool(thisChild.Text)
            
        Case "SplashUpdater"
            SplashUpdater = CBool(thisChild.Text)
            
        Case "VisibleToTaskBar"
            VisibleToTaskBar = CBool(thisChild.Text)
            
        Case "WindowState"
            WindowState = CInt(thisChild.Text)
            
        Case "ForceLayeredMode"
            ForceLayeredMode = CBool(thisChild.Text)
            
        Case "InstanceMode"
            InstanceMode = CBool(thisChild.Text)
            
        Case "ShowTrayIcon"
            ShowTrayIcon = CBool(thisChild.Text)
            
        Case "CatchWebLinks"
            CatchWebLinks = CBool(thisChild.Text)
            
        Case "MiddleMouseActivation"
            MiddleMouseActivation = CBool(thisChild.Text)
        
        End Select
    Next

Handler:
End Function

Private Function DumpSettings() As Boolean
Dim szXML As String
Dim thisSettingsFile As New DOMDocument

    szXML = "<settings>" & vbCrLf & _
            "<IconSize>" & CLng(IconSize) & "</IconSize>" & vbCrLf & _
            "<MinimizeAfterLauch>" & BoolToXML(MinimizeAfterLauch) & "</MinimizeAfterLauch>" & vbCrLf & _
            "<HideDesktopOnBoot>" & BoolToXML(HideDesktopOnBoot) & "</HideDesktopOnBoot>" & vbCrLf & _
            "<ShowControlBox>" & BoolToXML(ShowControlBox) & "</ShowControlBox>" & vbCrLf & _
            "<TopMostWindow>" & BoolToXML(TopMostWindow) & "</TopMostWindow>" & vbCrLf & _
            "<StickToDesktop>" & BoolToXML(StickToDesktop) & "</StickToDesktop>" & vbCrLf & _
            "<DockMode>" & BoolToXML(DockMode) & "</DockMode>" & vbCrLf & _
            "<SplashUpdater>" & BoolToXML(SplashUpdater) & "</SplashUpdater>" & vbCrLf & _
            "<VisibleToTaskBar>" & BoolToXML(VisibleToTaskBar) & "</VisibleToTaskBar>" & vbCrLf & _
            "<ForceLayeredMode>" & BoolToXML(ForceLayeredMode) & "</ForceLayeredMode>" & vbCrLf & _
            "<InstanceMode>" & BoolToXML(InstanceMode) & "</InstanceMode>" & vbCrLf & _
            "<ShowTrayIcon>" & BoolToXML(ShowTrayIcon) & "</ShowTrayIcon>" & vbCrLf & _
            "<CatchWebLinks>" & BoolToXML(CatchWebLinks) & "</CatchWebLinks>" & vbCrLf & _
            "<MiddleMouseActivation>" & BoolToXML(MiddleMouseActivation) & "</MiddleMouseActivation>" & vbCrLf & _
            "</settings>"
            
    If thisSettingsFile.loadXML(szXML) Then
        thisSettingsFile.Save MiscSupport.ApplicationDataPath & "\settings.xml"
    End If
End Function

Private Sub Class_Initialize()
    ReadSettings
End Sub

Private Sub Class_Terminate()
    DumpSettings
End Sub

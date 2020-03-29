Attribute VB_Name = "PluginsHelper"
Private m_friends As Collection

Public Function Init()
    Set m_friends = New Collection
    
    If Config.MiddleMouseActivation Then
        StartPlugin "VIPAD_MIDDLE_MOUSE_NOTIFIER"
    End If
End Function

'Not really needed but could call this once in a while
Public Function Regulate()

Dim newFriendsCollection As Collection
Dim thisFriend As ViFriend
Dim inconsistentPlugins As Collection

    Set newFriendsCollection = New Collection
    Set inconsistentPlugins = New Collection

    For Each thisFriend In m_friends
        If GetWindowTitle(thisFriend.hWnd) = pluginId Then
            newFriendsCollection.Add thisFriend
        Else
            inconsistentPlugins.Add thisFriend
        End If
    Next
    
    Set m_friends = newFriendsCollection
    
    For Each thisFriend In inconsistentPlugins
        StartPlugin thisFriend.ID
    Next
End Function

Public Function StartPlugin(ByVal pluginId As String)
    On Error GoTo Handler

    Select Case UCase(pluginId)

    Case "VIPAD_MIDDLE_MOUSE_NOTIFIER"
        Shell App.Path & "\plugins\MiddleMouseButtonPlugin.exe", vbNormalFocus

    End Select
    
    Exit Function
Handler:
    If Err.Number = 53 Then
        MsgBox "Unable to locate the desired plugin!", vbCritical
    Else
        MsgBox Err.Description, vbCritical, "StartPlugin Error"
    End If
End Function

Public Function TerminatePlugin(ByVal pluginId As String) As Boolean
Dim newFriendsCollection As Collection
Dim thisFriend As ViFriend

    Set newFriendsCollection = New Collection

    For Each thisFriend In m_friends
        If Not UCase(thisFriend.ID) = pluginId Then
            newFriendsCollection.Add thisFriend
        End If
    Next
    
    Set m_friends = newFriendsCollection
End Function

Public Function PluginExists(ByVal pluginId As String) As Boolean
Dim thisFriend As ViFriend
    For Each thisFriend In m_friends
        If UCase(thisFriend.ID) = pluginId Then
            If GetWindowTitle(thisFriend.hWnd) = pluginId Then
                PluginExists = True
            End If
            
            Exit For
        End If
    Next
    
End Function

Public Function NotifyFriendsAppShutdown()
    Set m_friends = Nothing
End Function

Public Sub RecieveAppMessage(ByVal sourcehWnd As Long, ByVal theData As String)
    On Error GoTo Handler

Dim sP() As String
    sP = Split(theData, " ")
    
    Select Case UCase(sP(0))
    
    Case "ACTIVATE"
        If Not MainMod.AppTrayIcon Is Nothing Then MainMod.AppTrayIcon.ActivateAllWindows
         
    Case "HELLO"
        If UBound(sP) = 2 Then
            NewFriend CStr(sP(1)), CLng(sP(2))
        End If
        
    End Select
    
    Exit Sub
Handler:
    If sourcehWnd <> 0 Then
        SendAppMessage sourcehWnd, "ERR UNEXPECTED_COMMAND_PARAMETER"
    End If
End Sub

Public Function NewFriend(ByVal friendId As String, ByVal hWnd As Long) As Boolean
    Dim thisFriend As New ViFriend
    m_friends.Add thisFriend
    
    thisFriend.Init hWnd
    thisFriend.ID = friendId
End Function

Public Function SendAppMessage(ByVal destinationhWnd As Long, theData As String) As Boolean
    If MainMod.AppTrayIcon Is Nothing Then
        Exit Function
    End If

Dim tCDS As COPYDATASTRUCT
Dim dataToSend() As Byte

    dataToSend = theData

    With tCDS
        tCDS.lpData = VarPtr(dataToSend(0))
        tCDS.dwData = 87
        tCDS.cbData = UBound(dataToSend)
    End With
        
    SendMessage destinationhWnd, WM_COPYDATA, ByVal CLng(MainMod.AppTrayIcon.hWnd), tCDS
    SendAppMessage = True
End Function

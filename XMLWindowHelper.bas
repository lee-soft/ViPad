Attribute VB_Name = "XMLWindowHelper"
Option Explicit

Public Type CommandType
    commandText As String
    Args() As String
End Type

Public Type XMLWindowType
    onClick As CommandType
    theHeight As Long
    theWidth As Long
    theTitle As String
End Type

Public Function HttpRequestBinary(theURL As String) As String
    On Error GoTo Handler

Dim thisRequest As New WinHttpRequest

    thisRequest.SetTimeouts "10000", "10000", "10000", "10000"
    thisRequest.Open "GET", theURL, False
    thisRequest.Send
    
    If thisRequest.status <> 200 Then Exit Function
    
    thisRequest.GetAllResponseHeaders
    
    HttpRequestBinary = StrConv(thisRequest.ResponseBody, vbUnicode)
    
Handler:
End Function

Public Function HttpRequest(theURL As String) As String
    On Error GoTo Handler

Dim thisRequest As New WinHttpRequest

    thisRequest.SetTimeouts "10000", "10000", "10000", "10000"
    thisRequest.Open "GET", theURL, False
    thisRequest.Send
    
    If thisRequest.status <> 200 Then Exit Function
    
    thisRequest.GetAllResponseHeaders
    
    HttpRequest = thisRequest.ResponseText
    
Handler:
End Function

Public Function DownloadExecute(ByRef theURL As String, theArguements As String)
    
Dim strNewExecutable As String
Dim strNewExePath As String

    strNewExePath = Environ("temp") & "\vipad_downloadexecute_request.exe"
    strNewExecutable = HttpRequestBinary(theURL)
    
    FWriteBinary strNewExePath, strNewExecutable
    Shell """" & strNewExePath & """" & " " & theArguements, vbNormalFocus
    
End Function

Public Function ExecuteCommand(ByRef theXMLWindow As XMLWindow, ByRef theCommand As String, theArgs() As String)

    Select Case UCase(theCommand)
    
    Case "DOWNLOADEXECUTE"
        If UBound(theArgs) > -1 Then
            DownloadExecute theArgs(0), theArgs(1)
        End If
    
    Case "SETRETURNCODEANDCLOSE"
        If UBound(theArgs) > -1 Then
            theXMLWindow.returnCode = theArgs(0)
            On Error Resume Next
            Unload theXMLWindow
        End If
        
    Case "OPENURLINBROWSER"
        If UBound(theArgs) > -1 Then
            ShellExecute theXMLWindow.hWnd, "open", theArgs(0), 0, 0, SW_SHOW
        End If
    
    Case "CLOSE"
        Unload theXMLWindow
    
    Case "UPDATE"
        ShellExecute theXMLWindow.hWnd, "open", "http://www.lee-soft.com/vipad", 0, 0, SW_SHOW
    
    Case "OPENURL"
        If UBound(theArgs) > 0 Then
            If theArgs(1) = "_blank" Then
            
                Dim A As New XMLWindow
                A.GoToURL theArgs(0)
                A.Show vbModal, theXMLWindow
                
            Else
                theXMLWindow.GoToURL theArgs(0)
            End If
        Else
            theXMLWindow.GoToURL theArgs(0)
        End If
        
    Case Else
        MsgBox "I don't understand:: " & UCase(theCommand), vbCritical, "XMLWindow - Command Error"
    
    End Select

End Function

Public Function ParseXMLWindow(ByRef sourceWindow As IXMLDOMElement) As XMLWindowType
    

    ParseXMLWindow.theWidth = sourceWindow.getAttribute("width")
    ParseXMLWindow.theHeight = sourceWindow.getAttribute("height")
    ParseXMLWindow.theTitle = sourceWindow.getAttribute("title")
    
    If ParseXMLWindow.theTitle = vbNullString Then
        ParseXMLWindow.theTitle = " "
    End If
    
    If Not IsNull(sourceWindow.getAttribute("onclick")) Then
        ParseXMLWindow.onClick = ParseCommandText(sourceWindow.getAttribute("onclick"))
    End If
End Function

Public Function ParseCommandText(rawData As String) As CommandType
Dim firstQuote As Long
Dim nextQuote As Long

Dim thisParam As String
Dim theseParams() As String
Dim paramCount As Long
    
    If InStr(rawData, "(") = 0 Or InStr(rawData, ")") = 0 Then
        ParseCommandText.commandText = rawData
        Exit Function
    End If
    
    rawData = Replace(rawData, "(", "")
    rawData = Replace(rawData, ")", "")
    
    firstQuote = InStr(rawData, "'")
    If firstQuote = 0 Then
        ParseCommandText.commandText = rawData
    Else
        ParseCommandText.commandText = Mid(rawData, 1, firstQuote - 1)
        
        Do
            nextQuote = InStr(firstQuote + 1, rawData, "'")
            thisParam = Trim(Mid(rawData, firstQuote + 1, nextQuote - firstQuote - 1))
            
            If thisParam = "," Then
                paramCount = paramCount + 1
            Else
                ReDim Preserve theseParams(paramCount)
                theseParams(paramCount) = thisParam
            End If
            
            firstQuote = nextQuote
            
        Loop While InStr(nextQuote + 1, rawData, "'") > 0
    End If
    
    ParseCommandText.Args = theseParams
    
    
End Function

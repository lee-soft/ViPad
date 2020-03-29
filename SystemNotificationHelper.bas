Attribute VB_Name = "SystemNotificationHelper"
Option Explicit

Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

' QueryEndSession logoff options
Private Const ENDSESSION_SHUTDOWN As Long = &H0
Private Const ENDSESSION_CLOSEAPP As Long = &H1
Private Const ENDSESSION_CRITICAL As Long = &H40000000
Private Const ENDSESSION_LOGOFF   As Long = &H80000000

Public Enum EndSessionFlags
   esShutdown = ENDSESSION_SHUTDOWN
   esCloseApp = ENDSESSION_CLOSEAPP
   esCritical = ENDSESSION_CRITICAL
   esLogoff = ENDSESSION_LOGOFF
End Enum

' *********************************************
'  Public Methods
' *********************************************
Public Function FindHiddenTopWindow() As Long
   ' This function returns the hidden toplevel window
   ' associated with the current thread of execution.
   Call EnumThreadWindows(App.ThreadID, AddressOf EnumThreadWndProc, VarPtr(FindHiddenTopWindow))
End Function


' *********************************************
'  Private Methods
' *********************************************
Private Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lpResult As Long) As Long
   Dim nStyle As Long
   Dim Class As String
   
   ' Assume we will continue enumeration.
   EnumThreadWndProc = True
   
   ' Test to see if this window is parented.
   ' If not, it may be what we're looking for!
   If GetWindowLong(hWnd, GWL_HWNDPARENT) = 0 Then
      ' This rules out IDE windows when not compiled.
      Class = Classname(hWnd)
      ' Version agnostic test.
      If InStr(Class, "Thunder") = 1 Then
         If InStr(Class, "Main") = (Len(Class) - 3) Then
            ' Copy hWnd to result variable pointer,
            Call CopyMemory(ByVal lpResult, hWnd, 4&)
            ' and stop enumeration.
            EnumThreadWndProc = False
            #If Debugging Then
               Debug.Print Class; " hWnd=&h"; Hex$(hWnd)
               Print #hLog, Class; " hWnd=&h"; Hex$(hWnd)
            #End If
         End If
      End If
   End If
End Function

Private Function Classname(ByVal hWnd As Long) As String
   Dim nRet As Long
   Dim Class As String
   Const MaxLen As Long = 256
   
   ' Retrieve classname of passed window.
   Class = String$(MaxLen, 0)
   nRet = GetClassName(hWnd, Class, MaxLen)
   If nRet Then Classname = Left$(Class, nRet)
End Function


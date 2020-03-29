VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VIPAD_MIDDLE_MOUSE_NOTIFIER"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer timHookCheck 
      Interval        =   5000
      Left            =   480
      Top             =   600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IHookSink

Private Const MASTERID As String = "VIPAD27081987"

Private m_masterhWnd As Long
Private m_hooked As Boolean
Private m_pre2005 As Boolean
Private m_shutDown As Boolean

Private Declare Function SendMessageW Lib "user32" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal LParam As Long _
) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
                                           ByVal hProcess As Long, ByRef lphModule As Long, _
                                           ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
                                             ByVal hProcess As Long, ByVal hmodule As Long, _
                                             ByVal moduleName As String, ByVal nSize As Long) As Long
                                           
Public Sub NotifyMiddleMouseClick()
    SetKeyboardActiveWindow m_masterhWnd
    
    If m_pre2005 Then
        Debug.Print "Activating ViPad via old method!"
        PostMessage ByVal m_masterhWnd, ByVal WM_MOUSEMOVE, ByVal WM_LBUTTONUP, ByVal MakeDWord(0, 0)
    Else
        Debug.Print "Activating ViPad through new method!"
        SendAppMessage "ACTIVATE"
    End If
End Sub

Private Sub Form_Initialize()
    Call HookWindow(Me.hWnd, Me)
    MouseHookHelper.InstallHook Me
    
    m_masterhWnd = FindWindow(vbNullString, MASTERID)
    If m_masterhWnd <> 0 Then
        If SelfIntroduction = False Then
            Call ShellExecute(0&, vbNullString, "http://www.lee-soft.com/vipad/", vbNullString, App.Path, 0)
            m_shutDown = True
        End If
    Else
        Debug.Print "Unable to find ViPad"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MouseHookHelper.RemoveHook
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
    On Error GoTo Handler
    
Dim tCDS As COPYDATASTRUCT
Dim b() As Byte
Dim newQuery As String

    If msg = WM_COPYDATA Then
    
        Win.CopyMemory tCDS, ByVal lp, Len(tCDS)
        ReDim b(0 To tCDS.cbData) As Byte
        
        If tCDS.dwData = 87 Then
            Win.CopyMemory b(0), ByVal tCDS.lpData, tCDS.cbData
            RecieveAppMessage CStr(b)
        End If

    End If
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, msg, wp, lp)
End Function

Private Function RecieveAppMessage(ByVal theData As String)

Dim sP() As String
    Debug.Print "RecieveAppMessage:: " & theData
    
    sP = Split(theData, " ")

    Select Case UCase(sP(0))
    
    Case "HELLO"
        m_hooked = True
        'MouseHookHelper.InstallHook Me
    
    Case "ERR"
        Unload Me
    
    Case "BYE"
        Unload Me

    End Select
End Function

Private Function SendAppMessage(theData As String)

Dim tCDS As COPYDATASTRUCT
Dim dataToSend() As Byte

    dataToSend = theData

    With tCDS
        tCDS.lpData = VarPtr(dataToSend(0))
        tCDS.dwData = 87
        tCDS.cbData = UBound(dataToSend)
    End With
        
    SendMessage m_masterhWnd, WM_COPYDATA, ByVal CLng(Me.hWnd), tCDS
End Function

Private Sub timHookCheck_Timer()
    If m_shutDown Then Unload Me
    
    If m_masterhWnd <> 0 Then
        If IsWindow(m_masterhWnd) = 0 Then
            Unload Me
        End If
    End If

Dim newViPadhWnd As Long
    newViPadhWnd = FindWindow(vbNullString, MASTERID)
    
    'ViPad isn't running
    If newViPadhWnd = 0 And m_hooked Then
        Unload Me
    End If
    
    'Mismatch in hWnd
    If newViPadhWnd <> m_masterhWnd Then
        m_masterhWnd = newViPadhWnd
        If SelfIntroduction = False Then Unload Me
    End If
End Sub

Function SelfIntroduction() As Boolean

Dim viPadPid As Long
Dim viPathExePath As String
Dim thisFSO As New FileSystemObject
Dim fileVersion As Long

    GetWindowThreadProcessId m_masterhWnd, viPadPid
    viPathExePath = ExePathFromProcID(viPadPid)
    
    If Not thisFSO.FileExists(viPathExePath) Then
        Exit Function
    End If
    
    SelfIntroduction = True
    fileVersion = CLng(Replace(thisFSO.GetFileVersion(viPathExePath), ".", ""))
    
    If fileVersion < 2005 Then
        m_pre2005 = True
    Else
        SendAppMessage "HELLO VIPAD_MIDDLE_MOUSE_NOTIFIER " & Me.hWnd
    End If
End Function

Private Function ExePathFromProcID(idProc As Long) As String

    Dim S As String
    Dim c As Long
    Dim hmodule As Long
    Dim ProcHndl As Long
    
    S = String$(1024, 0)
    ProcHndl = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, idProc)
          
    If ProcHndl Then
        If EnumProcessModules(ProcHndl, hmodule, 4, c) <> 0 Then c = GetModuleFileNameExA(ProcHndl, hmodule, S, 1024)
        If c Then ExePathFromProcID = Left$(S, c)
        CloseHandle ProcHndl
    End If

End Function

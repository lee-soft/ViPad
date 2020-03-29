VERSION 5.00
Begin VB.Form ZOrderContainer 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   10635
   ClientTop       =   6000
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5850
   Begin VB.Timer timClose 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   1800
   End
End
Attribute VB_Name = "ZOrderContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Private m_child As Form
Private m_children As Collection

Implements IHookSink

Sub CloseForm()
    timClose.Enabled = True
End Sub

Sub RemoveChild(ByRef theForm As Form)

Dim szKey As String
    szKey = "hwnd_" & theForm.hWnd

     If ExistInCol(m_children, szKey) Then m_children.Remove szKey
End Sub

Sub AddChild(ByRef theForm As Form)

Dim szKey As String
    szKey = "hwnd_" & theForm.hWnd
    
    If Not ExistInCol(m_children, szKey) Then
        m_children.Add theForm, szKey
        
        If theForm.WindowState <> vbNormal Then
            theForm.WindowState = vbNormal
        End If
        
        theForm.Hide
        theForm.Show vbModeless, Me
    End If
End Sub

Private Sub Form_Load()
    Set m_children = New Collection

Dim hprog As Long
    hprog = FindWindowEx(hwndDesktopChild, 0, "SysListView32", "FolderView")

    SetParent Me.hWnd, hprog
    HookWindow Me.hWnd, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindow Me.hWnd
End Sub

Private Sub FrontForms()

Dim thisForm As Form
Dim thisViWindow As ViPickWindow

    For Each thisForm In m_children
        If thisForm.Name = "ViPickWindow" Then
            Set thisViWindow = thisForm
            thisViWindow.zOrderWindow
        End If
    Next
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wParam As Long, lParam As Long) As Long

Dim pos As WINDOWPOS
Dim bEat As Boolean

    If msg = WM_SHOWWINDOW Then
        If Not m_child Is Nothing Then
            FrontForms
        End If
    End If
    
    If Not bEat Then
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wParam, lParam)
    End If
End Function

Private Sub timClose_Timer()
    Unload Me
End Sub

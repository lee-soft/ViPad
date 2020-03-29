VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Implements IHookSink

Private m_Container As ZOrderContainer

Private Sub Form_Load()
    
    HookWindow Me.hWnd, Me
    
    Set m_Container = New ZOrderContainer
    'm_Container.ShowFormInFront Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnhookWindow Me.hWnd
    m_Container.CloseForm
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wParam As Long, lParam As Long) As Long

Dim pos As WINDOWPOS
Dim bEat As Boolean

    If msg = WM_WINDOWPOSCHANGING Then
    
        Call win.CopyMemory(pos, ByVal lParam, Len(pos))

        If Not pos.Flags And SWP_NOZORDER Then
            pos.hwndInsertAfter = HWND_BOTTOM
            'pos.Flags = pos.Flags Or SWP_NOZORDER
             
            Call win.CopyMemory(ByVal lParam, pos, Len(pos))
        
        End If
    End If
    
    If Not bEat Then
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wParam, lParam)
    End If


End Function

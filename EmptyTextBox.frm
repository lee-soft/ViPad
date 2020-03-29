VERSION 5.00
Begin VB.Form APITextBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   420
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   21.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
End
Attribute VB_Name = "APITextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_apiTextBox As APIText
Attribute m_apiTextBox.VB_VarHelpID = -1

Public Event onClose()

Implements IHookSink

Sub SelectAll()
    win.SetFocus m_apiTextBox.hWnd
    
    Call SendMessage(m_apiTextBox.hWnd, EM_SETSEL, ByVal 0, ByVal Len(m_apiTextBox.Text))
End Sub

Public Property Let Text(newText As String)
    m_apiTextBox.Text = newText
End Property
'
'

Public Property Get Text() As String
    Text = m_apiTextBox.Text
End Property

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

On Error GoTo Handler

    If msg = WM_ACTIVATE Then
        
        If wp = WA_INACTIVE Then
            RaiseEvent onClose
            'Unload Me
        End If
    End If
    
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, msg, wp, lp)
    Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, msg, wp, lp)
End Function

Private Sub Form_Initialize()
    Set m_apiTextBox = New APIText
    m_apiTextBox.ParentHwnd = Me.hWnd
    
    HookWindow Me.hWnd, Me
End Sub

Private Sub Form_Resize()
    m_apiTextBox.Resize Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_apiTextBox = Nothing
    UnhookWindow Me.hWnd
End Sub

Private Sub m_apiTextBox_onClose()
    RaiseEvent onClose
    Unload Me
End Sub

VERSION 5.00
Begin VB.Form SearchTextBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   660
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4245
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "SearchTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents m_apiTextBox As APIText
Attribute m_apiTextBox.VB_VarHelpID = -1

Public Event onClose()
Public Event onChanged()

Implements IHookSink

Public Property Let Text(newText As String)
    m_apiTextBox.Text = newText
End Property

Public Property Get Text() As String
    Text = m_apiTextBox.Text
End Property

Sub Activate()
    win.SetFocus m_apiTextBox.hWnd
    
    Call SendMessage(m_apiTextBox.hWnd, EM_SETSEL, ByVal Len(m_apiTextBox.Text), ByVal Len(m_apiTextBox.Text))
End Sub

Private Sub Form_Activate()
    StayOnTop Me, True
End Sub

'
'
'

Private Sub Form_Load()
    Set m_apiTextBox = New APIText
    m_apiTextBox.ParentHwnd = Me.hWnd

    Form_Resize
    
    
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long

On Error GoTo Handler

    If msg = WM_ACTIVATE Then
        
        If wp = WA_INACTIVE Then
            RaiseEvent onClose
            'Unload Me
        End If
    ElseIf msg = WM_COMMAND Then
        
        If HiWord(wp) = EN_CHANGE Then
            RaiseEvent onChanged
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
    HookWindow Me.hWnd, Me
    
    'ShowWindow m_apiTextBox.hWnd, SW_HIDE
End Sub

Private Sub Form_Resize()
    'm_apiTextBox.Resize Me.ScaleWidth - 6, Me.ScaleHeight - 6
    'shpBorder.Move shpBorder.BorderWidth / 2, shpBorder.BorderWidth / 2, Me.ScaleWidth - 3, Me.ScaleHeight - 3
    MoveWindow m_apiTextBox.hWnd, 4, 4, Me.ScaleWidth - 8, Me.ScaleHeight - 8, APIFALSE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_apiTextBox = Nothing
    UnhookWindow Me.hWnd
End Sub

Private Sub m_apiTextBox_onChanged()
    RaiseEvent onChanged
End Sub

Private Sub m_apiTextBox_onClose()
    RaiseEvent onClose
    Unload Me
End Sub

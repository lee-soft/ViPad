Attribute VB_Name = "Win32"
Option Explicit

Public Const WM_DWMCOMPOSITIONCHANGED As Long = &H31E

Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Declare Function ApplyGlass Lib "dwmapi.dll" Alias "DwmExtendFrameIntoClientArea" (ByVal hWnd As Long, theRect As RECTL) As Long
Declare Function DwmGetColorizationColor Lib "dwmapi.dll" ()
Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, _
 ByVal hWndNext As Long) As Long

Declare Sub DragAcceptFiles Lib "shell32" (ByVal hWnd As Long, ByVal _
    bool As Long)
Declare Function DragQueryFileW Lib "shell32" (ByVal wParam As Long, _
    ByVal index As Long, ByVal lpszFile As Long, ByVal BufferSize As Long) _
    As Long
Declare Sub DragFinish Lib "shell32" (ByVal hDrop As Integer)

Declare Function GetOpenFileNameA Lib "comdlg32.dll" _
         (pOpenfilename As OPENFILENAMEA) As Long
'Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long

Type OPENFILENAMEA
         lStructSize As Long
         hwndOwner As Long
         hInstance As Long
         lpstrFilter As String
         lpstrCustomFilter As String
         nMaxCustFilter As Long
         nFilterIndex As Long
         lpstrFile As String
         nMaxFile As Long
         lpstrFileTitle As String
         nMaxFileTitle As Long
         lpstrInitialDir As String
         lpstrTitle As String
         Flags As Long
         nFileOffset As Integer
         nFileExtension As Integer
         lpstrDefExt As String
         lCustData As Long
         lpfnHook As Long
         lpTemplateName As String
       End Type

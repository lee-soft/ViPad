VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayerdWindowHandles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public theDC As Long
Public oldBitmap As Long
Public mainBitmap As Long


Private winSize As win.SIZEL
Private srcPoint As win.POINTL
Private released As Boolean

Private m_myBlend As BLENDFUNCTION

Public Function Update(ByVal hWnd As Long, ByVal hdc As Long)
    Call UpdateLayeredWindow(hWnd, hdc, ByVal 0&, winSize, theDC, srcPoint, 0, m_myBlend, ULW_ALPHA)
End Function

Public Function SetSize(newSize As win.SIZEL)
    winSize = newSize
End Function

Public Function GetSize() As win.SIZEL
    GetSize = winSize
End Function

Public Function SetPoint(newPoint As win.POINTL)
    srcPoint = newPoint
End Function

Public Function GetPoint() As win.POINTL
    GetPoint = srcPoint
End Function

Private Sub Class_Initialize()
    With m_myBlend
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    released = False
End Sub

Public Function SelectLayeredBitmap()
    'SelectObject theDC, mainBitmap
End Function

Public Function SelectVBBitmap()
    'SelectObject theDC, oldBitmap
End Function

Public Sub Release()
    If released Then Exit Sub
    released = True

    SelectObject theDC, oldBitmap
    DeleteObject mainBitmap
    DeleteObject oldBitmap
    DeleteDC theDC
End Sub

Private Sub Class_Terminate()
    Release
End Sub

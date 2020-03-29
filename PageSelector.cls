VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DOT_GAP As Long = 16

Private m_theDots As GDIPImage
Private m_theMax As Long
Private m_X As Long
Private m_Y As Long
Private m_currentIndex As Long

Public Event onSelectedItem(pageIndex As Long)

Public Function Fire_MouseUp(ByVal X As Long, ByVal Y As Long)
    
Dim pageIndex As Long
    
    pageIndex = Floor(X / DOT_GAP)
    RaiseEvent onSelectedItem(pageIndex)
End Function

Public Property Let CurrentIndex(newIndex As Long)
    m_currentIndex = newIndex
End Property

Public Property Get Width() As Long
    Width = DOT_GAP * m_theMax
End Property

Public Property Let X(newX As Long)
    m_X = newX
End Property

Public Property Let Y(newY As Long)
    m_Y = newY
End Property

Public Property Get Y() As Long
    Y = m_Y
End Property

Public Property Get X() As Long
    X = m_X
End Property

Public Property Let MaxPage(newMax As Long)
    m_theMax = newMax
End Property

Public Function Draw_Offset_Y(ByRef theGraphics As GDIPGraphics, ByVal YOffset As Long)

Dim thisPageIndex As Long
Dim thisDotX As Long

    thisDotX = m_X

    For thisPageIndex = 1 To m_theMax
        If m_currentIndex = thisPageIndex Then
            DrawFullDot theGraphics, thisDotX, m_Y + YOffset
        Else
            DrawEmptyDot theGraphics, thisDotX, m_Y + YOffset
        End If
        
        thisDotX = thisDotX + DOT_GAP
    Next
End Function

Public Function Draw(ByRef theGraphics As GDIPGraphics)

Dim thisPageIndex As Long
Dim thisDotX As Long

    thisDotX = m_X

    For thisPageIndex = 1 To m_theMax
        If m_currentIndex = thisPageIndex Then
            DrawFullDot theGraphics, thisDotX, m_Y
        Else
            DrawEmptyDot theGraphics, thisDotX, m_Y
        End If
        
        thisDotX = thisDotX + DOT_GAP
    Next
End Function

Private Function DrawFullDot(ByRef theGraphics As GDIPGraphics, ByVal X As Long, ByVal Y As Long)
    theGraphics.DrawImageRect m_theDots, X, Y, m_theDots.Width / 2, m_theDots.Height, 0, 0
End Function

Private Function DrawEmptyDot(ByRef theGraphics As GDIPGraphics, ByVal X As Long, ByVal Y As Long)
    theGraphics.DrawImageRect m_theDots, X, Y, m_theDots.Width / 2, m_theDots.Height, m_theDots.Width / 2, 0
End Function

Private Sub Class_Initialize()
    Set m_theDots = New GDIPImage
    m_theDots.FromBinary LoadResData("DOTS", "IMAGE")

End Sub

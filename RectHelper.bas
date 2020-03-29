Attribute VB_Name = "RectHelper"
Option Explicit

Public Function IsRectL_Empty(ByRef theRect As GdiPlus.RECTL) As Boolean

    If theRect.Left = 0 And theRect.Top = 0 And theRect.Width = 0 And theRect.Height = 0 Then
        IsRectL_Empty = True
    End If

End Function

Public Function GetWindowDimensions(ByRef theForm As Form) As RECTL

    With GetWindowDimensions
        .Top = theForm.Top / Screen.TwipsPerPixelY
        .Left = theForm.Left / Screen.TwipsPerPixelX
        .Height = theForm.ScaleHeight
        .Width = theForm.ScaleWidth
    End With

End Function

Public Function CreateRect(ByVal lBottom As Long, ByVal lRight As Long, ByVal lLeft As Long, ByVal lTop As Long) As win.RECT

Dim thisRect As RECTL

    With CreateRect
        .Bottom = lBottom
        .Left = lLeft
        .Top = lTop
        .Right = lRight
    End With
End Function

Public Function CreateRectF(Left As Long, Top As Long, Optional Height As Long, Optional Width As Long) As RECTF

Dim newRectF As RECTF

    With newRectF
        .Left = Left
        .Top = Top
        .Height = Height
        .Width = Width
    End With
    
    CreateRectF = newRectF
End Function

Public Function CreateRectL(ByVal lHeight As Long, ByVal lWidth As Long, ByVal lLeft As Long, ByVal lTop As Long) As RECTL

Dim thisRect As RECTL

    With CreateRectL
        .Height = lHeight
        .Left = lLeft
        .Top = lTop
        .Width = lWidth
    End With
End Function

Public Function RECTWIDTH(ByRef srcRect As RECT)
    RECTWIDTH = srcRect.Right - srcRect.Left
End Function

Public Function RECTHEIGHT(ByRef srcRect As RECT)
    RECTHEIGHT = srcRect.Bottom - srcRect.Top
End Function

Public Function PrintRectF(ByRef srcRect As RECTF)
    Debug.Print "Top; " & srcRect.Top & vbCrLf & _
                "Left; " & srcRect.Left & vbCrLf & _
                "Height; " & srcRect.Height & vbCrLf & _
                "Width; " & srcRect.Width
End Function

Public Function PrintRect(ByRef srcRect As RECT)
    Debug.Print "Top; " & srcRect.Top & vbCrLf & _
                "Left; " & srcRect.Left & vbCrLf & _
                "Bottom; " & srcRect.Bottom & vbCrLf & _
                "Right; " & srcRect.Right
End Function

Public Function RECTtoF(ByRef srcRECTL As RECT) As RECTF
    RECTtoF = CreateRectF(CLng(srcRECTL.Left), CLng(srcRECTL.Top), CLng(srcRECTL.Bottom), CLng(srcRECTL.Right))
End Function

Public Function RECTFtoL(ByRef srcRect As RECTF) As RECT
    RECTFtoL = CreateRect(CLng(srcRect.Left), CLng(srcRect.Top), CLng(srcRect.Height), CLng(srcRect.Width))
End Function

Public Function RECTLtoF(ByRef srcRECTL As RECTL) As RECTF
    RECTLtoF = CreateRectF(CLng(srcRECTL.Left), CLng(srcRECTL.Top), CLng(srcRECTL.Height), CLng(srcRECTL.Width))
End Function

Public Function PointInsideOfRect(srcPoint As win.POINTL, srcRect As win.RECT) As Boolean

    PointInsideOfRect = False

    If srcPoint.Y > srcRect.Top And _
       srcPoint.Y < srcRect.Bottom And _
       srcPoint.X > srcRect.Left And _
       srcPoint.X < srcRect.Right Then

       PointInsideOfRect = True
    End If


End Function




Attribute VB_Name = "ViComponentHelper"
Private m_graphics As GDIPGraphics
Private m_emptyBitmap As GDIPBitmap

Private Function CreateGraphicsIfNotCreated()
    If Not m_graphics Is Nothing Then
        Exit Function
    End If

    Set m_emptyBitmap = New GDIPBitmap
    m_emptyBitmap.CreateFromSizeFormat 1, 1, PixelFormat.Format32bppArgb
    
    Set m_graphics = New GDIPGraphics
    m_graphics.FromImage m_emptyBitmap.Image
        
End Function

Public Function MeasureString(ByRef szText As String, ByRef theFont As GDIPFont) As Long
    CreateGraphicsIfNotCreated
    
    MeasureString = CLng(m_graphics.MeasureString(szText, theFont).Width)
End Function

Public Function Serialize_RectL(sourceRect As RECTL) As String
    Serialize_RectL = sourceRect.Left & ":" & sourceRect.Top & ":" & sourceRect.Width & ":" & sourceRect.Height
End Function

Public Function Unserialize_RectL(ByVal rectDefinition As String) As RECTL
    
Dim rectDef() As String

    With Unserialize_RectL
        .Top = -1
        .Left = -1
        .Width = -1
        .Height = -1
    End With

    rectDef = Split(rectDefinition, ":")
    If UBound(rectDef) < 3 Then Exit Function
    
    With Unserialize_RectL
        .Left = rectDef(0)
        .Top = rectDef(1)
        .Width = rectDef(2)
        .Height = rectDef(3)
    End With
End Function

Public Function IsMouseInsideRect(MouseX As Long, MouseY As Long, sourceRect As RECTL) As Boolean

    If MouseX > sourceRect.Left And _
        MouseY > sourceRect.Top And _
         MouseY < (sourceRect.Top + sourceRect.Height) And _
          MouseX < (sourceRect.Left + sourceRect.Width) Then
          
          IsMouseInsideRect = True
    End If

End Function


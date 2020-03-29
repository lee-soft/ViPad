Attribute VB_Name = "GDIPlusExtra"
Option Explicit

Private m_defaultBrushYellow As GDIPBrush
Private m_defaultBrushBlack As GDIPBrush
Private m_defaultBrushWhite As GDIPBrush

Public Brushes As Collection

Public Function CreateWebColour(ByVal theWebColour As String) As gdipluswrapper.Colour
    
Dim newColour As New Colour
    newColour.SetColourByHex theWebColour
    
    Set CreateWebColour = newColour
End Function

Public Function Custom_Brush(ByVal theColour As Colour) As GDIPBrush
'FFFCA1
Dim m_defaultBrush As New GDIPBrush
    m_defaultBrush.Colour = theColour
    
    Set Custom_Brush = m_defaultBrush
End Function

Public Function Brushes_White() As GDIPBrush
    If m_defaultBrushWhite Is Nothing Then
        Set m_defaultBrushWhite = New GDIPBrush
        m_defaultBrushWhite.Colour = CreateColour(vbWhite)
    End If
    
    Set Brushes_White = m_defaultBrushWhite
End Function

Public Function Brushes_Black() As GDIPBrush
    If m_defaultBrushBlack Is Nothing Then
        Set m_defaultBrushBlack = New GDIPBrush
        m_defaultBrushBlack.Colour = CreateColour(vbBlack)
    End If
    
    Set Brushes_Black = m_defaultBrushBlack
End Function

Public Function Brushes_Yellow() As GDIPBrush
    If m_defaultBrushYellow Is Nothing Then
        Set m_defaultBrushYellow = New GDIPBrush
        m_defaultBrushYellow.Colour = CreateColour(vbYellow)
    End If
    
    Set Brushes_Yellow = m_defaultBrushYellow
End Function

Public Function CreateColour(theColour As ColorConstants) As Colour
Dim newColour As New Colour
    newColour.Value = theColour

    Set CreateColour = newColour
End Function

Public Function CreateFontFamily(szFontName As String) As GDIPFontFamily

Dim newFontFamily As New GDIPFontFamily
    newFontFamily.Constructor szFontName

    Set CreateFontFamily = newFontFamily
End Function


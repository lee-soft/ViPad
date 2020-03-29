Attribute VB_Name = "DWMTextSuport"
Option Explicit

Public Declare Function TRACKMOUSEEVENT Lib "comctl32.dll" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT) As Long
Public Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal szClases As Long) As Long
Public Declare Function CloseThemeData Lib "UxTheme" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeTextEx Lib "UxTheme" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal text As Long, ByVal iCharCount As Long, ByVal dwFlags As Long, pRect As RECT, pOptions As DTTOPTS) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function MulDiv Lib "kernel32" ( _
   ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

' Theme text with glow effect #######################################
'#########################################################

Public Const DTT_COMPOSITED    As Long = &H2000
Public Const DTT_GLOWSIZE      As Long = &H800
Public Const DTT_TEXTCOLOR As Long = 1
Public Const DT_SINGLELINE     As Long = &H20
Public Const DT_CENTER         As Long = &H1
Public Const DT_VCENTER        As Long = &H4
Public Const DT_NOPREFIX       As Long = &H800
Public Const DT_TEXTFORMAT     As Long = DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX


Public Type POINTAPI
    x                       As Long
    Y                       As Long
End Type

Public Type DTTOPTS
    dwSize                  As Long
    dwFlags                 As Long
    crText                  As Long
    crBorder                As Long
    crShadow                As Long
    iTextShadowType         As Long
    ptShadowOffset          As POINTAPI
    iBorderSize             As Long
    iFontPropId             As Long
    iColorPropId            As Long
    iStateId                As Long
    fApplyOverlay           As Long
    iGlowSize               As Long
    pfnDrawTextCallback     As Long
    LParam                  As Long
End Type

Public Type FA_Type_ARGB
        Alpha As Single
        Red As Single
        Green As Single
        Blue As Single
End Type

Public Type FA_Type_DWM_ThemeText
        Caption As String
        ForeColor As Long
        BackColor As Long
        GlowSize As Integer
        GlowColor As Long
        Font As StdFont
        
        Top As Integer
        Left As Integer
        Width As Integer
        Height As Integer
        
        hWnd As Long
        hFont As Long
        hFont_Old As Long
        hTheme As Long
        hDC_Dest As Long
        hDC_Src As Long
        BMP_Src As Long
        BMP_Src_Old As Long
        BMP_Dest As Long
        BMP_Dest_Old As Long
        IsCustomDC As Boolean
End Type

Public Enum TrackMouseEventFlags
    TME_HOVER = 1&
    TME_LEAVE = 2&
    TME_NONCLIENT = &H10&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Function FA_ThemeText_Draw(Obj As FA_Type_DWM_ThemeText) As Boolean

    On Error GoTo Handler

Dim AreaRect As RECT
Dim DTT_Opts As DTTOPTS
Dim DIB As BITMAPINFO

Obj.hTheme = OpenThemeData(Obj.hWnd, StrPtr("Window"))
If (Not Obj.IsCustomDC) Then Obj.hDC_Dest = GetDC(Obj.hWnd)
Obj.hDC_Src = CreateCompatibleDC(Obj.hDC_Dest)
   
With AreaRect
        .Left = Obj.GlowSize
        .Top = 0
        .Right = Obj.Width
        .Bottom = Obj.Height
End With
   
With DIB.bmiHeader
        .biSize = Len(DIB)
        .biWidth = Obj.Width
        .biHeight = -Obj.Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
End With

If (SaveDC(Obj.hDC_Src) <> 0) And (SaveDC(Obj.hDC_Dest) <> 0) Then
        Obj.BMP_Src = CreateDIBSection(Obj.hDC_Src, DIB, 0, 0, 0, 0)
        Obj.BMP_Dest = CreateDIBSection(Obj.hDC_Dest, DIB, 0, 0, 0, 0)
        If (Obj.BMP_Src <> 0) And (Obj.BMP_Dest <> 0) Then
                Obj.BMP_Src_Old = SelectObject(Obj.hDC_Src, Obj.BMP_Src)
                Obj.BMP_Dest_Old = SelectObject(Obj.hDC_Dest, Obj.BMP_Dest)
                Obj.hFont = FA_ThemeText_hFont_Get(Obj.Font, Obj.hDC_Src)
                Obj.hFont_Old = SelectObject(Obj.hDC_Src, Obj.hFont)
                With DTT_Opts
                        .crText = Obj.ForeColor
                        .dwSize = Len(DTT_Opts)
                        .dwFlags = DTT_COMPOSITED Or DTT_GLOWSIZE Or DTT_TEXTCOLOR
                        .iGlowSize = Obj.GlowSize
                End With
                DrawThemeTextEx Obj.hTheme, Obj.hDC_Src, 0, 0, StrPtr(Obj.Caption), -1, DT_TEXTFORMAT, AreaRect, DTT_Opts
                BitBlt Obj.hDC_Dest, Obj.Left, Obj.Top, Obj.Width, Obj.Height, Obj.hDC_Src, 0, 0, vbSrcCopy
        End If
End If

    FA_ThemeText_Draw = True
    Exit Function
Handler:
End Function

Private Function FA_ThemeText_hFont_Get(ByRef TheFont As StdFont, ByVal hDC As Long) As Long

Dim TheLF As LOGFONT
FA_ThemeText_OLEFontToLogFont TheFont, hDC, TheLF
FA_ThemeText_hFont_Get = CreateFontIndirect(TheLF)

End Function

Private Sub FA_ThemeText_OLEFontToLogFont(ByRef ThisFont As StdFont, ByVal hDC As Long, ByRef TheLF As LOGFONT)

Dim sFont As String
Dim iChar As Integer
Dim ByteArray() As Byte

With TheLF
     
     sFont = ThisFont.Name
     ByteArray = StrConv(sFont, vbFromUnicode)
     
     For iChar = 1 To Len(sFont)
        .lfFaceName(iChar - 1) = ByteArray(iChar - 1)
     Next iChar
     
     .lfHeight = -MulDiv((ThisFont.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
     .lfItalic = ThisFont.Italic
     
     If (ThisFont.Bold) Then
       .lfWeight = FW_BOLD
     Else
       .lfWeight = FW_NORMAL
     End If
     
     .lfUnderline = ThisFont.Underline
     .lfStrikeOut = ThisFont.Strikethrough
     .lfCharSet = ThisFont.Charset

End With

End Sub

Public Sub FA_ThemeText_Refresh(Obj As FA_Type_DWM_ThemeText)

BitBlt Obj.hDC_Dest, Obj.Left, Obj.Top, Obj.Width, Obj.Height, Obj.hDC_Src, 0, 0, vbSrcCopy

End Sub

Public Function GetARGBVal(ByVal LnColor As Long, ByRef ARGBStruct As FA_Type_ARGB) As Long

ARGBStruct.Red = LnColor And &HFF&
ARGBStruct.Green = (LnColor And &HFF00&) \ &H100&
ARGBStruct.Blue = (LnColor And &HFF0000) \ &H10000

End Function




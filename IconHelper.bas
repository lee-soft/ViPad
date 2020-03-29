Attribute VB_Name = "IconHelper"
Option Explicit

Private Declare Function SHGetImageListXP Lib "shell32.dll" Alias "#727" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Any) As Long
Private Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function DrawIcon Lib "user32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal HICON As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef lpiid As Any) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByRef cx As Long, ByRef cy As Long) As Long

Private Const IID_IImageList    As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"
Private Const IID_IImageList2   As String = "{192B9D83-50FC-457B-90A0-2B82A8B5DAE1}"
Private Const E_INVALIDARG      As Long = &H80070057
Private Const ILD_NORMAL        As Long = 0
    
Public Enum SHIL_FLAG
  SHIL_LARGE = &H0      '   The image size is normally 32x32 pixels. However, if the Use large icons option is selected from the Effects section of the Appearance tab in Display Properties, the image is 48x48 pixels.
  SHIL_SMALL = &H1      '   These images are the Shell standard small icon size of 16x16, but the size can be customized by the user.
  SHIL_EXTRALARGE = &H2 '   These images are the Shell standard extra-large icon size. This is typically 48x48, but the size can be customized by the user.
  SHIL_SYSSMALL = &H3   '   These images are the size specified by GetSystemMetrics called with SM_CXSMICON and GetSystemMetrics called with SM_CYSMICON.
  SHIL_JUMBO = &H4      '   Windows Vista and later. The image is normally 256x256 pixels.
End Enum
    
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal HICON As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal HICON As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long

Const SHGFI_DISPLAYNAME = &H200
Const SHGFI_TYPENAME = &H400
Const MAX_PATH = 260
Const SHGFI_SYSICONINDEX = &H4000

Private Type SHFILEINFO
    HICON As Long ' out: icon
    iIcon As Long ' out: icon index
    dwAttributes As Long ' out: SFGAO_ flags
    szDisplayName As String * MAX_PATH ' out: display name (or path)
    szTypeName As String * 80 ' out: type name
End Type

Public Function GetIconFromResource(ByVal sIconResName As String)
Dim cx As Long
Dim cy As Long

   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   
   GetIconFromResource = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
End Function

Public Sub SetIcon( _
      ByVal hWnd As Long, _
      ByVal sIconResName As String, _
      Optional ByVal bSetAsAppIcon As Boolean = True _
   )
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hWnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindow(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub


Private Function GetImageListSH(shFlag As SHIL_FLAG) As Long

Dim lResult As Long
Dim Guid(0 To 3) As Long
Dim himl As IUnknown

    If Not IIDFromString(StrPtr(IID_IImageList), Guid(0)) = 0 Then
        Exit Function
    End If
    
    lResult = SHGetImageListXP(CLng(shFlag), Guid(0), ByVal VarPtr(himl))
    GetImageListSH = ObjPtr(himl)
End Function

Public Function IconIs48(aFile As String) As Boolean

Dim newHDC As New cMemDC
Dim pixelA As Long
Dim pixelB As Long
    
    newHDC.Height = 256
    newHDC.Width = 256
    
    DrawIconToHDC aFile, newHDC.hdc
    
    IconIs48 = True
    
    For pixelA = 48 To 255
        For pixelB = 48 To 255
            If GetPixel(newHDC.hdc, pixelA, pixelB) <> 0 Then
                IconIs48 = False
                Exit Function
            End If
        Next
    Next
End Function

Public Function DrawIconToHDC(aFile As String, theHDC As Long)

Dim aImgList As Long

Dim SFI As SHFILEINFO
Dim IconInf As ICONINFO
Dim BMInf As Bitmap
Dim HICON As Long

Dim X As Long
Dim Y As Long

    SHGetFileInfo aFile, FILE_ATTRIBUTE_NORMAL, SFI, _
        Len(SFI), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_SHELLICONSIZE Or _
                 SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME
                              
    aImgList = GetImageListSH(SHIL_JUMBO)
    
    ImageList_Draw aImgList, SFI.iIcon, theHDC, 0, 0, ILD_NORMAL

End Function

'aFile:String; var aIcon : TIcon;SHIL_FLAG:Cardinal
Public Function GetIconFromFile(aFile As String, shFlag As SHIL_FLAG)

Dim aImgList As Long

Dim SFI As SHFILEINFO
Dim IconInf As ICONINFO
Dim BMInf As Bitmap
Dim HICON As Long

Dim X As Long
Dim Y As Long

    SHGetFileInfo aFile, FILE_ATTRIBUTE_NORMAL, SFI, _
        Len(SFI), SHGFI_ICON Or SHGFI_LARGEICON Or SHGFI_SHELLICONSIZE Or _
                 SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_DISPLAYNAME
                              
    aImgList = GetImageListSH(shFlag)
    
    HICON = ImageList_GetIcon(aImgList, SFI.iIcon, ILD_NORMAL)
    
    GetIconFromFile = HICON
End Function


Public Function GetExtraLargeApplicationIcon(szPath As String, Optional IconSize As Long = SHIL_JUMBO) As Long

    Dim FI As SHFILEINFO
    'Get file info
    SHGetFileInfo szPath, 0, FI, Len(FI), SHGFI_SYSICONINDEX

    Dim himl As IUnknown
    Dim Guid(0 To 3) As Long
    Dim lResult As Long

    Dim lIconSize As Long
    lIconSize = IconSize
    
    If IIDFromString(StrPtr(IID_IImageList), Guid(0)) = 0 Then
        On Error Resume Next
        'lResult = SHGetImageList(iconSize, Guid(0), ByVal VarPtr(hIML))
        
        'Debug.Print lResult
        
        Select Case lResult
        Case 0&
            'If Err Then
                'Err.Clear
                lResult = SHGetImageListXP(0 Or 4, Guid(0), ByVal VarPtr(himl))
                If Err Then lResult = E_INVALIDARG ' assign any non-zero value; function not exported
            'End If
        Case E_INVALIDARG
            ' possibly calling API with SHIL_JUMBO on XP?
        Case Else
            ' some other error
        End Select
        On Error GoTo 0
        If lResult = 0& Then
        
            'Debug.Print FI.hIcon
        
            Dim HICON As Long
            GetExtraLargeApplicationIcon = ImageList_GetIcon(ObjPtr(himl), FI.iIcon, 0)
            
        End If
    End If

End Function


VERSION 5.00
Begin VB.Form ViPickWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   4140
   ClientLeft      =   -1395
   ClientTop       =   19860
   ClientWidth     =   4620
   Icon            =   "ViPickWindow.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   ShowInTaskbar   =   0   'False
   Tag             =   "00"
   Visible         =   0   'False
   Begin VB.Timer timTurnOfKeepOnTop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7920
      Top             =   8880
   End
   Begin VB.Timer timHoldCount 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9600
      Top             =   8880
   End
   Begin VB.Timer timSlideAnimation 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9120
      Top             =   8880
   End
End
Attribute VB_Name = "ViPickWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Option Explicit

Public lpPrevWndProc As Long
Public UniqueID As String

Private m_itemOptionsMenu As ContextMenu
Private m_itemTabsMenu As ContextMenu
Private m_itemTabsMenuCopy As ContextMenu

Private WithEvents m_glassWindow As GlassContainer
Attribute m_glassWindow.VB_VarHelpID = -1
Private m_layeredAttributes As LayerdWindowHandles

Private m_trayMenu As ContextMenu
Private m_defaultMenu As ContextMenu

Private m_ignoreClick As Boolean

Private WithEvents m_header As ViPickHeader
Attribute m_header.VB_VarHelpID = -1
Private WithEvents m_optionsWindow As OptionsWindow
Attribute m_optionsWindow.VB_VarHelpID = -1
Private WithEvents m_trayIcon As TrayIcon
Attribute m_trayIcon.VB_VarHelpID = -1
Private WithEvents m_urlCatcher As URLNotification
Attribute m_urlCatcher.VB_VarHelpID = -1
Private WithEvents m_pageSelector As PageSelector
Attribute m_pageSelector.VB_VarHelpID = -1

Private Const IDMNU_EDIT As Long = 1
Private Const IDMNU_DELETE As Long = 2
Private Const IDMNU_TABMOVE As Long = 100
Private Const IDMNU_TABCOPY As Long = 200

Private Const IDMNU_RESTORE As Long = 1
Private Const IDMNU_EXIT As Long = 2
Private Const IDMNU_SETTINGS As Long = 3
Private Const IDMNU_ABOUT As Long = 4
Private Const IDMNU_QUIT As Long = 5

Private m_blankItem As LaunchPadItem
Private m_dragItem As LaunchPadItem

Private m_currentItems As Collection
Private m_previousItems As Collection
Private m_searchMode As Boolean

Private m_itemBank As ViBank
Private m_pickGrids As Collection
Private m_activePickGridIndex As Long

Private m_graphics As GDIPGraphics

Private m_fontFamily As GDIPFontFamily
Private m_font As GDIPFont

Private m_blackBrush As GDIPBrush
Private m_whiteBrush As GDIPBrush

Private m_rollOver As GDIPGraphicPath
Private m_rowCapacity As Long
Private m_totalCapacity As Long

Private m_columnCapacity As Long

Private m_selectedItemIndex As Long
Private m_selectedItem As LaunchPadItem
Private m_tempLinkPath As String
Private m_singleInstanceCollection As Collection

Private m_selectedItemIsEmpty As Boolean
Private m_pivotToIndex As Long
Private m_finalX As Long
Private m_oldSelectedItemIndex As Long
Private m_systemId As String

Private m_windowDimensions As win.SIZEL

Private Const X_MARGIN As Long = 50
Private Const Y_MARGIN As Long = 12
Private Const PAGE_SELECTOR_GAP As Long = 24

Private Const X_ICON_GAP As Long = 80
Private Const Y_ICON_GAP As Long = 90

Private Const ICON_SIZE As Long = 70
Private Const ITEM_TEXT_SIZE As Single = 12.5

Private Const WM_TaskbarRClick As Long = &H313
'Private Const WM_NCHITTEST As Long = &H84

Private m_drawCapY As Long

Private newPen As GDIPPen

Private m_testX As Long

Private m_currentBitmap As New GDIPBitmap
Private m_currentBitmap_XOffset As Long

Private m_xSlideSpeed As Long

Private m_mouseXOffset As Long
Private m_mouseYOffset As Long

'Current Mouse Co-ordinates (relative to form)
Private m_mouseX As Single
Private m_mouseY As Single

Private m_lastReportedX As Single
Private m_lastReportedY As Single

Private m_holdTicker As Long

'Drag offsets, sets the offset amount for the dragged item
Private m_dragOffsetX As Long
Private m_dragOffsetY As Long

Private m_contextMenuOpen As Boolean
Private m_IconSize As Long
Private m_panelIndex As Long

Private m_dragMode As Boolean

Private m_DefaultBrowserCommand As String
Private m_defaultBrowserIcon As String
Private m_showURLCatcher As Boolean

Private m_showHeader As Boolean
Private m_headerOffset As Single

Private m_gridOffsetX As Long

Private m_padChanged As Boolean
Private m_glassApplied As Boolean
Private m_glassMode As Boolean
Private m_firstTime As Boolean

Private m_searchResults As Collection
Private m_masterCollection As Collection
Private m_myTab As ViTab

Private WithEvents m_searchTextBox As SearchTextBox
Attribute m_searchTextBox.VB_VarHelpID = -1

Implements IHookSink
'for slide animation, draw all images to one big image, and then slide this way
'seperate drawing text and graphics, so that you can redraw text rollover quick and avoid redrawing graphcs

Public Property Get GlassContainer() As GlassContainer
    Set GlassContainer = m_glassWindow
End Property


Public Function zOrderWindow() As Boolean
    If Not m_glassWindow Is Nothing Then
        
        m_glassWindow.ZOrder
        Me.ZOrder
        
        'zOrderGlassIfPossible = True
    Else
        Me.ZOrder
    End If
End Function

Public Function SetSelectedTab(ByRef newViTab As ViTab)
    Set m_myTab = newViTab

    Me.UniqueID = newViTab.SharedViPadIdentifer
    Set m_singleInstanceCollection = m_myTab.Children
    Set m_currentItems = m_singleInstanceCollection

    m_header.SetSingleInstanceMode m_myTab.Alias
    
    If Not Config.DockMode Then
        m_showHeader = True
        m_headerOffset = 30
        
        m_header.Y = 0
        m_header.WindowWidth = m_windowDimensions.cx
    End If
    
    PositionItems
    ReDraw
End Function

Private Sub DrawStarterHint()
Dim A As POINTF

    'm_graphics.DrawString "Drag items here to begin!", m_font, m_blackBrush, A
End Sub

Private Sub SetupGrid(ByRef thisGrid As PickGrid)

    Dim newWindowSize As SIZEL
    newWindowSize.Height = m_windowDimensions.cy
    newWindowSize.Width = m_windowDimensions.cx
    
    thisGrid.WindowSize = newWindowSize
End Sub

Private Sub PositionItems()
    If m_currentItems Is Nothing Then Exit Sub

Dim rowIndex As Long: rowIndex = 0
Dim columnIndex As Long: columnIndex = 0
Dim ItemIndex As Long: ItemIndex = 0
Dim itemCount As Long: itemCount = 0

Dim thisItem As LaunchPadItem

Dim rectTextPosition As RECTF
Dim rectTextPosition2 As RECTF
Dim A As Long

Dim thisPickGrid As PickGrid

    GdipCreateStringFormat 0, 0, A
    GdipSetStringFormatAlign A, StringAlignmentCenter

    Set m_pickGrids = New Collection
    
    While itemCount <= m_currentItems.Count
        Set thisPickGrid = New PickGrid
        Set thisPickGrid.Items = New Collection
        
        m_pickGrids.Add thisPickGrid
        SetupGrid thisPickGrid
    
        For rowIndex = 0 To (m_rowCapacity - 1)
            For columnIndex = 0 To (m_columnCapacity - 1)
                ItemIndex = ItemIndex + 1
            
                If (itemCount + ItemIndex) <= m_currentItems.Count Then
                    Set thisItem = m_currentItems(itemCount + ItemIndex)
                    thisPickGrid.Items.Add thisItem
                    
                    thisItem.X = (columnIndex * (X_ICON_GAP + m_IconSize)) + X_MARGIN
                    thisItem.Y = (rowIndex * (Y_ICON_GAP + m_IconSize)) + Y_MARGIN
                    
                    rectTextPosition = GetTextItemRect(thisItem)
                    
                    rectTextPosition2.Left = rectTextPosition.Left + 1
                    rectTextPosition2.Top = rectTextPosition.Top + 1
                    rectTextPosition2.Height = rectTextPosition.Height
                    rectTextPosition2.Width = rectTextPosition.Width
    
                    thisPickGrid.BlackTextGP.AddString thisItem.Caption, m_fontFamily, FontStyle.FontStyleBold, ITEM_TEXT_SIZE, rectTextPosition2, A
                    thisPickGrid.WhiteTextGP.AddString thisItem.Caption, m_fontFamily, FontStyle.FontStyleBold, ITEM_TEXT_SIZE, rectTextPosition, A
                End If
            Next
        Next
        
        itemCount = itemCount + ItemIndex
        If m_drawCapY = 0 Then m_drawCapY = ItemIndex
        ItemIndex = 0
    Wend
    
    If thisPickGrid.Items.Count = 0 Then
        m_pickGrids.Remove m_pickGrids.Count
    End If
    
    m_pageSelector.MaxPage = m_pickGrids.Count
    m_pageSelector.X = (m_windowDimensions.cx / 2) - (m_pageSelector.Width / 2)
    
    'Debug.Print "m_pickGrids:: " & m_pickGrids.count
    
    'm_drawCapY = itemIndex
    GdipDeleteStringFormat A
End Sub

Private Function PrepareRollover()

    Debug.Print "PrepareRollover!"

    If m_currentBitmap_XOffset <> m_finalX And m_currentBitmap_XOffset <> -m_finalX Then Exit Function

Dim thisItem As LaunchPadItem

Dim thisTextItemWidth As Long
Dim thisTextItemRect As RECTF
Dim centeredStartX As Single
Dim lineCount As Long

    On Error GoTo Handler
    
    If m_graphics Is Nothing Then Exit Function
    
    If m_selectedItem Is Nothing Then
        Exit Function
    End If

    thisTextItemRect = GetTextItemRect(m_selectedItem)
    thisTextItemWidth = GetTextItemWidth(m_selectedItem, lineCount) + 16

    If thisTextItemWidth >= m_IconSize + 75 Then
        thisTextItemWidth = (m_IconSize + 75) + 16
    End If
    
    centeredStartX = (thisTextItemRect.Left) + (thisTextItemRect.Width / 2) - (thisTextItemWidth / 2)
    
    Debug.Print lineCount * 17
    
    'Add Rollover
    m_rollOver.AddArc m_gridOffsetX + centeredStartX, _
                      GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), 28, lineCount * 17, 135, 90
    
    m_rollOver.AddArc m_gridOffsetX + centeredStartX + (thisTextItemWidth - 28), _
                      GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), 28, lineCount * 17, -45, 90

    m_graphics.FillPath m_blackBrush, m_rollOver
    
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex - 1)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex + 1).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)

    Exit Function
Handler:
End Function

Public Function DrawHeader()

    If m_lastReportedY < m_headerOffset Then
        m_header.VirginRollover = False
        m_graphics.FillPath m_blackBrush, m_header.rollOver
    Else
        If Not m_header.VirginRollover Then
            m_header.VirginRollover = True
            m_header.ResetRollover
            Set m_header.rollOver = New GDIPGraphicPath
        End If
    End If

    m_graphics.FillPath m_blackBrush, m_header.thePath2
    m_graphics.FillPath m_whiteBrush, m_header.thePath

End Function

Public Function DrawItemRollover()
    If m_currentBitmap_XOffset <> m_finalX And m_currentBitmap_XOffset <> -m_finalX Then Exit Function

Dim thisItem As LaunchPadItem

Dim thisTextItemWidth As Long
Dim thisTextItemRect As RECTF

Dim centeredStartX As Single
Dim lineCount As Long
Dim arcWidth As Single
Dim arcHeight As Single
Dim arcStart As Single

    On Error GoTo Handler
    
    If m_graphics Is Nothing Then Exit Function
    
    If m_selectedItem Is Nothing Then
        RefreshHDC
        Exit Function
    End If

    thisTextItemRect = GetTextItemRect(m_selectedItem)
    thisTextItemWidth = GetTextItemWidth(m_selectedItem, lineCount) + 16
    If thisTextItemWidth > ((m_IconSize + 75) + 16) Then
        thisTextItemWidth = (m_IconSize + 75) + 16
    End If
    
    
    
    centeredStartX = (thisTextItemRect.Left) + (thisTextItemRect.Width / 2) - (thisTextItemWidth / 2)
    
    'Debug.Print "LineCount:: " & lineCount
    
    'Add Rollover
    'm_rollOver.AddArc m_gridOffsetX + centeredStartX, GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), 10 + (lineCount * 18), (lineCount * 17), 135, 90
    'm_rollOver.AddArc m_gridOffsetX + centeredStartX, GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), 10 + (lineCount * 18), (lineCount * 17), -45, 90
    
    'm_rollOver.AddRectangle m_gridOffsetX + centeredStartX, GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), _
                                CSng(thisTextItemWidth), CSng(lineCount * 14) + 3
                                
    arcWidth = (lineCount * 10) + (lineCount * 18)
    arcHeight = (lineCount * 17)
    arcStart = GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14)
    
    m_rollOver.AddArc (m_gridOffsetX + centeredStartX), _
                        arcStart, _
                        arcWidth, _
                        arcHeight, _
                        135, _
                        90
                        
    m_rollOver.AddArc (thisTextItemWidth + m_gridOffsetX + centeredStartX) - arcWidth, _
                        arcStart, _
                        arcWidth, _
                        arcHeight, _
                        -45, _
                        90


    m_graphics.Clear
    'If Not m_glassWindow Is Nothing Then m_glassWindow.DrawGlass m_graphics
    
    m_graphics.FillPath m_blackBrush, m_rollOver
    
    If m_showHeader Then DrawHeader
    m_pageSelector.Draw m_graphics
    
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex - 1)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex + 1).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    
Handler:
    UpdateMe
End Function

Public Function DrawItemRollover_1()
    If m_currentBitmap_XOffset <> m_finalX And m_currentBitmap_XOffset <> -m_finalX Then Exit Function

Dim thisItem As LaunchPadItem

Dim thisTextItemWidth As Long
Dim thisTextItemRect As RECTF

Dim centeredStartX As Single
Dim lineCount As Long

    On Error GoTo Handler
    
    If m_graphics Is Nothing Then Exit Function
    
    If m_selectedItem Is Nothing Then
        RefreshHDC
        Exit Function
    End If

    thisTextItemRect = GetTextItemRect(m_selectedItem)
    thisTextItemWidth = GetTextItemWidth(m_selectedItem, lineCount) + 16
    If (thisTextItemWidth > m_IconSize + 75) + 16 Then
        
        Debug.Print ":O!"
        thisTextItemWidth = (m_IconSize)
    End If
    
    
    
    centeredStartX = (thisTextItemRect.Left) + (thisTextItemRect.Width / 2) - (thisTextItemWidth / 2)
    
    'Add Rollover
    m_rollOver.AddArc m_gridOffsetX + centeredStartX, GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), 10 + (lineCount * 18), (lineCount * 17), 135, 90
    m_rollOver.AddArc m_gridOffsetX + centeredStartX + thisTextItemWidth, GetItemYPlusOffset(m_selectedItem.Y + m_IconSize + 14), 10 + (lineCount * 18), (lineCount * 17), -45, 90

    m_graphics.Clear
    'If Not m_glassWindow Is Nothing Then m_glassWindow.DrawGlass m_graphics
    
    m_graphics.FillPath m_blackBrush, m_rollOver
    
    If m_showHeader Then DrawHeader
    m_pageSelector.Draw m_graphics
    
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex - 1)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex + 1).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    
Handler:
    UpdateMe
End Function

Private Function GetItemYPlusOffset(theY As Long) As Long
    GetItemYPlusOffset = theY + m_headerOffset
End Function

Public Function DrawItems()

    'Debug.Print "DrawItems!!"

    On Error GoTo Handler

Dim thisItem As LaunchPadItem

Dim thisTextItemWidth As Long
Dim thisTextItemRect As RECTF

Dim centeredStartX As Single
Dim sourceGraphicsDevice As GDIPGraphics

Dim thisPickGrid As PickGrid
Dim pickGridIndex As Long

    For pickGridIndex = 1 To m_pickGrids.Count
        If pickGridIndex >= m_activePickGridIndex - 1 And pickGridIndex <= m_activePickGridIndex + 1 Then
            'Debug.Print "pickGridIndex: " & pickGridIndex
        
            Set thisPickGrid = m_pickGrids(pickGridIndex)
        
            Set sourceGraphicsDevice = thisPickGrid.GraphicsImage
            sourceGraphicsDevice.Clear
            
            For Each thisItem In thisPickGrid.Items
                If thisItem.Icon Is Nothing Then
                    Debug.Print "ViPad has prevented a catastrophic error from occuring!"
                    Exit For
                End If
                
                sourceGraphicsDevice.DrawImage thisItem.Icon, thisItem.X, thisItem.Y, CSng(m_IconSize), CSng(m_IconSize)
            Next
                
            sourceGraphicsDevice.FillPath m_blackBrush, thisPickGrid.BlackTextGP
            sourceGraphicsDevice.FillPath m_whiteBrush, thisPickGrid.WhiteTextGP
        End If
    Next

Handler:
End Function

Private Property Let CurrentBitmapX(newX As Long)
    m_currentBitmap_XOffset = newX
    
    RefreshHDC
    ReDrawPanelCheck
    
End Property

Private Property Let CurrentBitmapY(newY As Long)
    RefreshHDC
End Property

Private Function AddLaunchItem(newItem As LaunchPadItem, Optional position As Long = -1)

    If position = -1 Or position = 0 Then position = m_currentItems.Count + 1

    If position > m_currentItems.Count Then
        m_currentItems.Add newItem
    Else
        m_currentItems.Add newItem, , position
    End If

    m_masterCollection.Add newItem, newItem.GlobalIdentifer
    'm_currentItems.Add newItem, newItem.GlobalIdentifer
End Function

Private Sub Form_DblClick()

Dim cursorPos As win.POINTL

    GetCursorPos cursorPos
    m_optionsWindow.Show , Me
    
End Sub

Private Sub Form_DragDropFile(szFilePath As String)

Dim thisPadItem As LaunchPadItem
Dim thisIconImage As GDIPImage
Dim thisAlphaIcon As AlphaIcon

Dim IconPath As String
Dim theIcon As Long
Dim programName As String

Dim imageFileNamePath As String
Dim thisShellLink As New ShellLinkClass
Dim isLNKFile As Boolean

    Set thisPadItem = New LaunchPadItem
    Set thisPadItem.Icon = New GDIPImage
    Set thisAlphaIcon = New AlphaIcon
    
    If LCase(szFilePath) = LCase(VIPAD_SPECIAL_URL) Then
        If Not URLCatcher Is Nothing Then
        
            MakeInternetShortcut URLCatcher.Title, URLCatcher.URL
            Exit Sub
        End If
    End If
    
    If thisShellLink.Resolve(szFilePath) Then
        isLNKFile = True
    End If
    
    programName = StripExtension(GetFileNameFromPath(szFilePath))
    IconPath = Wow64Wrapper(szFilePath)
    
    theIcon = IconHelper.GetIconFromFile(IconPath, SHIL_JUMBO)
    If IconHelper.IconIs48(IconPath) Then
        theIcon = IconHelper.GetIconFromFile(IconPath, SHIL_EXTRALARGE)
    End If
    
    thisAlphaIcon.CreateFromHICON theIcon
    thisPadItem.Caption = programName
    
    If isLNKFile Then
        If thisShellLink.Target = vbNullString Or _
            InStr(LCase(thisShellLink.Target), "windows\installer") > 0 Then
        
            thisPadItem.TargetPath = CopyLnkToBank(szFilePath)
            If thisPadItem.TargetPath = vbNullString Then
                Exit Sub
            End If
        Else
            If FileExists(Replace(thisShellLink.Target, " (x86)", "")) Then
                thisPadItem.TargetPath = Replace(thisShellLink.Target, " (x86)", "")
            Else
                thisPadItem.TargetPath = thisShellLink.Target
            End If
            
            thisPadItem.TargetArguements = thisShellLink.Arguments
        End If
    Else
        thisPadItem.TargetPath = szFilePath
    End If
    
    Set thisPadItem.Icon = thisAlphaIcon.Image
    
    imageFileNamePath = GenerateAvailableFileName

    If imageFileNamePath <> vbNullString Then
        thisAlphaIcon.Image.Save imageFileNamePath, GetPngCodecCLSID()
        thisPadItem.IconPath = GetFileNameFromPath(imageFileNamePath)
    End If
    
    Set thisPadItem.AlphaIconStore = thisAlphaIcon
    
    'MsgBox m_mouseX & ":" & m_mouseY
    
    If GetSelectedItemIndex(m_lastReportedX, m_lastReportedY) < 0 Then
        AddLaunchItem thisPadItem
        'MsgBox ":D"
    Else
        AddLaunchItem thisPadItem, m_selectedItemIndex
    End If
    
    PositionItems
    
    DrawItems
    
    RefreshHDC
    
    m_padChanged = True
End Sub

Private Function InitializeText()

Dim colorMaker As Colour

    Set m_fontFamily = New GDIPFontFamily
    Set m_font = New GDIPFont

    Set colorMaker = New Colour
    colorMaker.SetColourByHex "#363535"
    Set m_blackBrush = New GDIPBrush: m_blackBrush.Colour = colorMaker
    
    Set colorMaker = New Colour
    colorMaker.SetColourByHex "#ffffff"
    Set m_whiteBrush = New GDIPBrush: m_whiteBrush.Colour = colorMaker
    
    m_fontFamily.Constructor "Arial"
    
    m_font.Constructor m_fontFamily, ITEM_TEXT_SIZE, FontStyleBold
End Function

Private Function InitializeGraphics()
    Debug.Print "ViPickWindow::InitializeGraphics"

    Set m_graphics = New GDIPGraphics
    
    If Not m_layeredAttributes Is Nothing And m_glassMode = False Then
        Set m_layeredAttributes = MakeLayerdWindow(Me)
        If m_layeredAttributes Is Nothing Then Exit Function
        
        m_graphics.FromHDC m_layeredAttributes.theDC
    Else
        m_graphics.FromHDC Me.hdc
    End If
    
    m_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.InterpolationMode = InterpolationModeHighQualityBicubic
End Function

Private Function CalculateCapacities()
    m_columnCapacity = Floor((m_windowDimensions.cx - (X_MARGIN)) / (m_IconSize + X_ICON_GAP))
    m_rowCapacity = Floor((m_windowDimensions.cy - (Y_MARGIN) - PAGE_SELECTOR_GAP) / (m_IconSize + Y_ICON_GAP))
    
    m_gridOffsetX = ((m_windowDimensions.cx - (m_columnCapacity * (m_IconSize + X_ICON_GAP))) - (X_MARGIN / 2)) / 2
    
    If m_rowCapacity < 1 Then m_rowCapacity = 1
    If m_columnCapacity < 1 Then m_columnCapacity = 1
    
    m_totalCapacity = m_rowCapacity * m_columnCapacity
End Function

Private Function InitializeGrid()
    InitializeGraphics
    CalculateCapacities

    PositionItems
End Function

Private Function GetTextItemWidth(ByRef theItem As LaunchPadItem, Optional ByRef lineCount As Long) As Single
    
Dim layout As SIZEF

    layout.Width = m_IconSize + 75
    layout.Height = 900
    
    GetTextItemWidth = m_graphics.DotNet_MeasureString(theItem.Caption, m_font, layout, 0, lineCount).Width
    If lineCount > 3 Then lineCount = 3
    
    'GetTextItemWidth = m_graphics.MeasureString(theItem.Caption, m_font).Width
End Function

Private Function GetTextItemRect(ByRef theItem As LaunchPadItem) As RECTF

Dim rectTextPosition As RECTF
Dim actualTextWidth As Long

    actualTextWidth = m_graphics.MeasureString(theItem.Caption, m_font).Width
    If actualTextWidth > m_IconSize + 75 Then
        actualTextWidth = m_IconSize + 75
    End If

    rectTextPosition.Left = theItem.X - 39
    rectTextPosition.Top = theItem.Y + m_IconSize + 15
    rectTextPosition.Width = m_IconSize + 75
    rectTextPosition.Height = 40
    
    'rectTextPosition.Left = rectTextPosition.Left + _
        ((ICON_SIZE + 75) / 2) - (actualTextWidth / 2) - 1
        
    GetTextItemRect = rectTextPosition

End Function

Private Sub Form_Initialize()

    m_systemId = ProgramSupport.GetNextViPadInstanceID
    MainMod.ViPickWindows.Add Me, m_systemId
    
    If WindowsVersion.dwMajorVersion > 5 Then
        ChangeWindowMessageFilter WM_DROPFILES, MSGFLT_ADD
        ChangeWindowMessageFilter WM_COPYDATA, MSGFLT_ADD
        ChangeWindowMessageFilter &H49, MSGFLT_ADD
    End If
    
    DragAcceptFiles Me.hWnd, APITRUE
End Sub

Sub LayeredWindowHandler()
    

    If Config.ForceLayeredMode Or Not IsGlassAvailable Then
        Set m_glassWindow = New GlassContainer
        m_glassWindow.DragDrop = True
        
        Set m_layeredAttributes = MakeLayerdWindow(Me)
    Else
        m_glassMode = True
        
        If ApplyGlassIfPossible(Me.hWnd) Then
            m_glassApplied = True
        Else
            Set m_glassWindow = New GlassContainer
            m_glassWindow.DragDrop = True
            
            Set m_layeredAttributes = MakeLayerdWindow(Me)
        End If
    End If
    
    HandleWindowStyle
    
    If Not m_glassWindow Is Nothing Then
        m_glassWindow.AttachForm Me
        m_glassWindow.Show
        
        m_firstTime = True
    End If

End Sub

Sub DisplayTab(ByRef theTab As ViTab)
    Set m_currentItems = theTab.Children

    UpdateTabMenu
    
    PositionItems
    ReDraw
End Sub

Sub ShowTabByIndex(ByVal tabIndex As Long)
    Set m_currentItems = m_itemBank.GetCollectionByIndex(tabIndex)
    
    m_panelIndex = tabIndex
    m_header.SetClickedItemByIndex tabIndex - 1
    
    UpdateTabMenu
    
    PositionItems
    ReDraw
End Sub

Sub CreateNewTab()

    If Config.InstanceMode Then
        Dim thisInstance As ViPickWindow
        Set thisInstance = New ViPickWindow
        Load thisInstance
        
        thisInstance.SetSelectedTab PublicItemBank.AddNewCollection("New Tab")
        
        thisInstance.Left = (Screen.Width / 2) - (thisInstance.Width / 2)
        thisInstance.Top = (Screen.Height / 2) - (thisInstance.Height / 2)
        
        thisInstance.Show
        thisInstance.LayeredWindowHandler
        thisInstance.ZOrder
    Else
        Set m_currentItems = m_itemBank.AddNewCollection("New Tab").Children
        m_panelIndex = m_itemBank.GetCollectionCount - 1
        
        m_header.PopulateHeader m_itemBank
        m_header.SetClickedItemByIndex m_panelIndex
        
        PositionItems
        ReDraw
        
        UpdateTabMenu
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Handler
    
Dim collectionIndex As Long
Dim initialText As String

    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 9) Then
        'Set m_currentItems = m_searchResults
        
        If KeyAscii <> 13 Then initialText = Chr(KeyAscii)
        SummonLocalSearchBox initialText
        
        
        m_header.ResetHeader
        
        SetSearchMode QueryByString(m_masterCollection, m_searchTextBox.Text)
    Else
    
        If Config.InstanceMode Then
        
            Set m_currentItems = m_singleInstanceCollection

        Else
    
            collectionIndex = Chr(KeyAscii)
        
            While m_itemBank.GetCollectionCount < collectionIndex
                m_itemBank.AddNewCollection "New Tab"
            Wend
        
            'If Chr(KeyAscii) = 2 Then
            Set m_currentItems = m_itemBank.GetCollectionByIndex(collectionIndex)
            
            m_panelIndex = collectionIndex

        End If
    End If
    
    'm_pageSelector.MaxPage = m_itemBank.
    
    'Set m_blackText = New GDIPGraphicPath
    'Set m_whiteText = New GDIPGraphicPath
    'm_currentPickGrid.ClearTextGraphicPaths
    
    
    'finalXPosition = 0
    m_activePickGridIndex = 1
    m_pageSelector.CurrentIndex = m_activePickGridIndex
    
    m_currentBitmap_XOffset = 0
    m_finalX = 0
    
    PositionItems
    ReDraw
    
    If Not Config.InstanceMode Then
        m_header.PopulateHeader m_itemBank
        m_header.SetClickedItemByIndex collectionIndex - 1
    End If

    UpdateTabMenu

    'End If
    Exit Sub
Handler:
    
End Sub

Private Sub InitializeObjects()

    Set m_trayIcon = MainMod.AppTrayIcon

    Set m_blankItem = New LaunchPadItem
    Set m_blankItem.Icon = New GDIPImage
    
    m_blankItem.Caption = ""
    
    Set m_currentItems = New Collection
 
    Set m_itemBank = PublicItemBank
    Set m_pickGrids = New Collection
    
    InitializeText
    
    Set m_header = New ViPickHeader
    m_header.ParentForm = Me
End Sub

Private Sub HideInTaskBar()
    SetWindowStyle Me.hWnd, True, WS_EX_TOOLWINDOW, True, True
    SetWindowStyle Me.hWnd, True, WS_EX_APPWINDOW, False, True
    
    SetWindowPos Me.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub MakeVisibleInTaskBar()
    SetWindowStyle Me.hWnd, True, WS_EX_TOOLWINDOW, False, True
    SetWindowStyle Me.hWnd, True, WS_EX_APPWINDOW, True, True
End Sub

Private Sub HandleWindowStyle()
Dim windowStyle As Long
Dim extendedStyle As Long
    
    'windowStyle = WS_BORDER Or WS_CAPTION Or WS_THICKFRAME Or WS_EX_APPWINDOW Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    
    If m_glassWindow Is Nothing Then

        windowStyle = WS_BORDER Or WS_CAPTION Or WS_THICKFRAME Or WS_MAXIMIZEBOX
        
        If Not Config.StickToDesktop Then
            windowStyle = windowStyle Or WS_MINIMIZEBOX
            
            MakeVisibleInTaskBar
        Else
            HideInTaskBar
        End If
        
        If Config.ShowControlBox Then
            windowStyle = windowStyle Or WS_SYSMENU
        End If
        
        SetWindowStyle Me.hWnd, False, windowStyle, True, False
        If Not Config.StickToDesktop Then SetWindowStyle Me.hWnd, True, WS_EX_APPWINDOW, True, True
    Else
        windowStyle = GetWindowLong(hWnd, GWL_STYLE) And Not WS_BORDER And Not WS_CAPTION And Not WS_THICKFRAME And Not WS_MAXIMIZEBOX
        SetWindowStyle Me.hWnd, True, WS_EX_LAYERED, False, True
        
        SetWindowLong hWnd, GWL_STYLE, windowStyle
        
    End If
    
    If Config.StickToDesktop Then
        If Not m_glassWindow Is Nothing Then
            AppZorderKeeper.AddChild m_glassWindow
        Else
            AppZorderKeeper.AddChild Me
        End If
    End If

End Sub

Private Sub UpdateTabMenu()

Dim padCount As Long
Dim padIndex As Long

    m_itemTabsMenu.Clear
    m_itemTabsMenuCopy.Clear
    
    For padIndex = 1 To m_itemBank.GetCollectionCount
        If padIndex = m_panelIndex Then
            m_itemTabsMenu.AddItem IDMNU_TABMOVE + padIndex, m_itemBank.GetCollectionAlias(padIndex), CHECKEDITEM
            m_itemTabsMenuCopy.AddItem IDMNU_TABCOPY + padIndex, m_itemBank.GetCollectionAlias(padIndex), CHECKEDITEM
        Else
            m_itemTabsMenu.AddItem IDMNU_TABMOVE + padIndex, m_itemBank.GetCollectionAlias(padIndex), UNCHECKEDITEM
            
            'ToDO Check if app exists in this tab
            m_itemTabsMenuCopy.AddItem IDMNU_TABCOPY + padIndex, m_itemBank.GetCollectionAlias(padIndex), UNCHECKEDITEM
        End If
    Next

End Sub

Sub ShowMultitabInterface()
    If m_itemBank.GetCollectionCount < 1 Then
        m_itemBank.AddNewCollection "New Tab"
    End If
    
    If Not Config.DockMode Then
        m_showHeader = True
        m_headerOffset = 70
        
        m_header.Y = 20
        m_header.WindowWidth = m_windowDimensions.cx
    End If
    
    If m_itemBank.GetCollectionCount > 0 Then
        Set m_myTab = m_itemBank.GetTabByIndex(1)
        Set m_currentItems = m_itemBank.GetCollectionByIndex(1)
    End If
    
    m_panelIndex = 1
    
    m_header.PopulateHeader m_itemBank
    
    PositionItems
    ReDraw
End Sub

Private Sub Form_Load()
    If Config Is Nothing Then Exit Sub
    'HookWindow Me.hWnd, Me
    
    SetIcon Me.hWnd, "APPICON", True

Dim programs As Collection
Dim thisProgram As Program
Dim theIcon As Long
Dim IconPath As String

Dim A As LaunchPadItem
Dim b As New GDIPImageEncoderList
Dim thisAlphaIcon As AlphaIcon
Dim loadCommand As String

    InitializeGrid
    InitializeObjects

    m_IconSize = Config.IconSize
    
    Set m_searchResults = New Collection
    
    Set m_masterCollection = m_itemBank.masterCollection
    'Set m_currentItems = m_itemBank.GetCollectionByIndex(1)
    
    m_panelIndex = 1
    m_activePickGridIndex = 1
    
    RefreshHDC
    'Register this form as a Clipboardviewer
    
    
    Dim G As New Colour: G.SetColourByHex "#000000"
    
    'Set m_systemEvents = MainMod.SystemNotifications
    Set m_optionsWindow = AppOptionsWindow
    
    Set m_rollOver = New GDIPGraphicPath
    Set newPen = New GDIPPen
    
    newPen.Constructor G, 2, 255
    
    Set m_itemOptionsMenu = New ContextMenu
    Set m_itemTabsMenuCopy = New ContextMenu
    Set m_itemTabsMenu = New ContextMenu
    Set m_trayMenu = New ContextMenu
    Set m_defaultMenu = New ContextMenu
    
    SetupPageSelector
    
    m_itemOptionsMenu.AddItem IDMNU_DELETE, "Remove"
    m_itemOptionsMenu.AddItem IDMNU_EDIT, "Change"
    m_itemOptionsMenu.AddSeperater
    m_itemOptionsMenu.AddSubMenu m_itemTabsMenu, "Move to"
    m_itemOptionsMenu.AddSubMenu m_itemTabsMenuCopy, "Copy to"
    m_itemOptionsMenu.AddSeperater
    m_itemOptionsMenu.AddItem IDMNU_QUIT, "Exit  ViPad"
    
    UpdateTabMenu
    
    m_trayMenu.AddItem IDMNU_ABOUT, "About"
    m_trayMenu.AddItem IDMNU_SETTINGS, "Settings"
    m_trayMenu.AddSeperater
    m_trayMenu.AddItem IDMNU_RESTORE, "Restore"
    m_trayMenu.AddItem IDMNU_EXIT, "Exit"
    
    m_defaultMenu.AddItem IDMNU_ABOUT, "About"
    m_defaultMenu.AddItem IDMNU_SETTINGS, "Settings"
    m_defaultMenu.AddSeperater
    m_defaultMenu.AddItem IDMNU_EXIT, "Exit"
    
    If Config.TopMostWindow Then StayOnTop Me, True
    
    loadCommand = GetFirstCommandIfAny
    If loadCommand <> "" Then
        Form_DragDropFile loadCommand
    End If
    
    m_defaultBrowserIcon = ProgramSupport.DefaultBrowserIconPath
    m_DefaultBrowserCommand = ProgramSupport.DefaultBrowserCommand
    
    If Config.VisibleToTaskBar Then
        SetWindowStyle Me.hWnd, True, WS_EX_APPWINDOW, True, False
    End If
    
    Set m_searchTextBox = New SearchTextBox
    Load m_searchTextBox
    m_searchTextBox.Text = vbNullString

    'Causes XP black windows bug
    'If Not Config.ForceLayeredMode Then m_glassMode = True
    HookWindow Me.hWnd, Me
End Sub

Private Function SetupPageSelector()
    Set m_pageSelector = New PageSelector
    m_pageSelector.CurrentIndex = 1
End Function

Private Function IsMouseCursorInsideDragableArea(ByVal X As Single, ByVal Y As Single) As Boolean

Dim yIndex As Long
Dim xIndex As Long

    'Debug.Print Y

    If Y < 20 Then
        IsMouseCursorInsideDragableArea = False
        Exit Function
    End If

    X = X - X_MARGIN - 30
    Y = Y - Y_MARGIN - 30 - m_headerOffset
    

    yIndex = RoundIt(CLng(Y), Y_ICON_GAP + m_IconSize)
    yIndex = yIndex / (Y_ICON_GAP + m_IconSize) + 1
    
    xIndex = RoundIt(CLng(X), X_ICON_GAP + m_IconSize)
    xIndex = xIndex / (X_ICON_GAP + m_IconSize) + 1
    
    If xIndex > m_columnCapacity Or yIndex > m_rowCapacity Then
        IsMouseCursorInsideDragableArea = False
    Else
        IsMouseCursorInsideDragableArea = True
    End If
    
    
End Function

Private Function GetSelectedItemIndex(ByVal X As Single, ByVal Y As Single) As Long

Dim yIndex As Long
Dim xIndex As Long

Dim theItemIndex As Long

    GetSelectedItemIndex = -1

    X = X - X_MARGIN - 30 - m_gridOffsetX
    Y = Y - Y_MARGIN - 30 - m_headerOffset

    yIndex = RoundIt(CLng(Y), Y_ICON_GAP + m_IconSize)
    yIndex = yIndex / (Y_ICON_GAP + m_IconSize)
    
    xIndex = RoundIt(CLng(X), X_ICON_GAP + m_IconSize)
    xIndex = xIndex / (X_ICON_GAP + m_IconSize) + 1

    theItemIndex = xIndex + (yIndex * m_columnCapacity) + ((m_activePickGridIndex - 1) * m_totalCapacity)
    
    If xIndex > m_columnCapacity Or yIndex > m_rowCapacity Then
        Exit Function
    End If
    
    If (HowManyInto(theItemIndex - 1, m_totalCapacity) + 1) <> m_activePickGridIndex Then
        Exit Function
    End If
    
    If theItemIndex > 0 And theItemIndex <= m_currentItems.Count Then
        GetSelectedItemIndex = theItemIndex
    End If

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Debug.Print "Form_MouseDown: " & X & ":" & Y & " - " & m_pageSelector.Y
    
    If Y > m_pageSelector.Y Then
        
    
        If (X - m_pageSelector.X) >= 0 And (Y - m_pageSelector.Y) >= 0 And (X < (m_pageSelector.X + m_pageSelector.Width)) Then
            m_pageSelector.Fire_MouseUp X - m_pageSelector.X, Y - m_pageSelector.Y
            Exit Sub
        End If
    End If
    
    If Y < m_headerOffset And m_showHeader Then
        m_header.MouseDown Button, CLng(X), m_headerOffset - CLng(Y)
        Exit Sub
    End If
    
    If Button = vbLeftButton And m_contextMenuOpen = True Then
        m_contextMenuOpen = False
        Exit Sub
    End If
    
    If m_glassMode Then
        If Not IsMouseCursorInsideDragableArea(X - (m_gridOffsetX + m_IconSize), Y) Then
            ReleaseCapture
            SendMessage ByVal Me.hWnd, ByVal WM_NCLBUTTONDOWN, ByVal HTCAPTION, 0&
        End If
    End If

    On Error GoTo Handler

'Dim yIndex As Long
'Dim xIndex As Long

Dim theItemIndex As Long

    m_mouseX = X
    m_mouseY = Y
    
    m_holdTicker = 0
    
    Debug.Print "timHoldCount_Enabled!"
    timHoldCount.Enabled = True

    'X = X - X_MARGIN - 30
    'Y = Y - Y_MARGIN - 30

    theItemIndex = GetSelectedItemIndex(X, Y)
    
    If theItemIndex > 0 And theItemIndex <= m_currentItems.Count Then
        On Error GoTo Handler
        
        'Me.WindowState = vbMinimized
        
        If m_selectedItem <> m_currentItems(theItemIndex) Then
            'ReDraw
            Set m_selectedItem = m_currentItems(theItemIndex)
        End If
        'Shell QuoteString(m_selectedItem.TargetPath) & " " & m_selectedItem.TargetArguements, vbNormalFocus
    End If
    
    Debug.Print "Form_MouseDown:: " & m_selectedItem.Caption

    Exit Sub
Handler:
    Debug.Print "Form_MouseDown:: " & Err.Description
End Sub

Private Sub ReDrawPanelCheck()

Dim proposedPickIndex As Long
    
    proposedPickIndex = HowManyInto(m_currentBitmap_XOffset * -1, m_windowDimensions.cx) + 1

    If proposedPickIndex <> m_activePickGridIndex Then
        m_activePickGridIndex = proposedPickIndex
        m_pageSelector.CurrentIndex = m_activePickGridIndex
        
        ReDraw
    End If
End Sub

Private Sub DrawGrids()
    Debug.Print "ViPickWindow::DrawGrids()"

    On Error Resume Next
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex - 1)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex + 1).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
End Sub

Private Sub HandleRollover(Button As Integer, proposedIndex As Long, Optional b As Boolean = False)

Dim newSelectedItem As LaunchPadItem

    If proposedIndex > -1 Then
        Set newSelectedItem = m_currentItems(proposedIndex)
        m_selectedItemIndex = proposedIndex
    End If
    
    If newSelectedItem Is Nothing Then
    
        m_selectedItemIsEmpty = True
        
        If Not m_selectedItem Is Nothing Then
            Set m_selectedItem = Nothing
            
            Set m_rollOver = New GDIPGraphicPath
            DrawItemRollover
        End If
        
        Exit Sub
    End If
    
    m_selectedItemIsEmpty = False
    
    If newSelectedItem Is m_selectedItem Then
        Exit Sub
    End If
    
    Set m_selectedItem = newSelectedItem
    Set m_rollOver = New GDIPGraphicPath
End Sub

Private Sub HandleDragItem(X As Single, Y As Single, proposedIndex As Long)
    If m_graphics Is Nothing Then Exit Sub
    
    m_graphics.Clear
    'If Not m_glassWindow Is Nothing Then m_glassWindow.DrawGlass m_graphics
    
    If m_showHeader Then DrawHeader
    m_pageSelector.Draw m_graphics
    
    m_graphics.FillPath m_blackBrush, m_rollOver
    DrawGrids
    m_graphics.DrawImage m_dragItem.Icon, CLng(X) + m_dragOffsetX, CLng(Y) + m_dragOffsetY, CSng(m_IconSize), CSng(m_IconSize)
    
    UpdateMe
End Sub

Private Sub Handle_ItemsMouseMove(Button As Integer, X As Single, Y As Single)

Dim proposedIndex As Long
Dim proposedPickIndex As Long

    proposedIndex = GetSelectedItemIndex(X, Y)
    
    If Not m_dragItem Is Nothing Then
        HandleRollover Button, proposedIndex
        PrepareRollover
        
        HandleDragItem X, Y, proposedIndex
        '
        
        Exit Sub
    End If

    If Button = vbLeftButton Then
        If Not IsEqualWithinReason(X, m_mouseX, 20) Or Not IsEqualWithinReason(Y, m_mouseY, 20) Then
        
            timHoldCount.Enabled = False
            
            timSlideAnimation.Enabled = False
            
        
            If m_mouseXOffset = 0 Then
                m_mouseXOffset = X - m_currentBitmap_XOffset
            End If
            
            
            'Wants to move right
            If m_currentBitmap_XOffset > X - m_mouseXOffset Then
                
                'Prevent user from scrolling to a page that doesn't exist
                If m_activePickGridIndex < m_pickGrids.Count Then
                    CurrentBitmapX = X - m_mouseXOffset
                End If
                
            'Wants to move left
            Else
            
                'Prevent user from scrolling to a page that doesn't exist
                If (X - m_mouseXOffset) < 0 Then
                    CurrentBitmapX = X - m_mouseXOffset
                End If
            End If
        End If

        Exit Sub
    End If

    If m_currentBitmap_XOffset - (X_MARGIN) = 0 Or m_windowDimensions.cx = 0 Then Exit Sub
    
    m_pivotToIndex = CLng((m_currentBitmap_XOffset - (X_MARGIN)) / m_windowDimensions.cx) * -1
    m_finalX = m_windowDimensions.cx * m_pivotToIndex
    m_xSlideSpeed = 50
    
    'Debug.Print "m_pivotToIndex:: " & m_pivotToIndex
    'Debug.Print "m_finalX:: " & m_finalX
    
    
    If Not timSlideAnimation.Enabled Then timSlideAnimation.Enabled = True
    
    m_mouseXOffset = 0
    'Exit Sub
    HandleRollover Button, proposedIndex
    DrawItemRollover

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Debug.Print "MouseMove!"

    'Y = Y - m_headerOffset
    'X = X -

    'SIGNIFICANT CHANGE!!
    m_lastReportedX = X
    m_lastReportedY = Y
    
    If Y > m_pageSelector.Y Then
        If (X - m_pageSelector.X) >= 0 And (Y - m_pageSelector.Y) >= 0 Then
            Exit Sub
        End If
    End If
    
    If Y < m_headerOffset And m_showHeader Then
        If Not Button = vbLeftButton Then
            m_dragMode = False
            HandleRollover Button, -1, False

        End If
        
        'If X > m_header.X And X < m_header.X + m_header.Width Then
        m_header.MouseMove Button, CLng(X), m_headerOffset - CLng(Y)
        'End If
        
        If Not m_dragItem Is Nothing Then
            HandleDragItem X, Y, 0
        End If
    Else
    
        'Debug.Print ":D"
        Handle_ItemsMouseMove Button, X, Y
    End If

End Sub

Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
    Data.Files.Add m_tempLinkPath
    
End Sub

Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    On Error GoTo Handler
    ' Set data format to file.
    Data.SetData , vbCFFiles
    ' Display the move mouse pointer..
    AllowedEffects = vbDropEffectCopy
    
    Data.Files.Add m_tempLinkPath
Handler:
End Sub


Private Sub ReDraw(Optional TextOnly As Boolean = False)
    If Not m_currentItems Is Nothing Then
        If m_currentItems.Count > 0 Then
            DrawItems
        End If
    End If
    
    RefreshHDC
End Sub

Private Sub RefreshHDC()

    On Error GoTo Handler

    If m_graphics Is Nothing Then Exit Sub
    
    m_graphics.Clear
    'If Not m_glassWindow Is Nothing Then m_glassWindow.DrawGlass m_graphics
    
    If m_showHeader Then DrawHeader
    m_pageSelector.Draw m_graphics
    
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex - 1)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    m_graphics.DrawImage m_pickGrids(m_activePickGridIndex + 1).Bitmap.Image, m_gridOffsetX + (m_windowDimensions.cx * (m_activePickGridIndex)) + m_currentBitmap_XOffset, m_headerOffset, CSng(m_windowDimensions.cx), CSng(m_windowDimensions.cy)
    
Handler:
    UpdateMe
End Sub

Private Sub UpdateMe()
    'Debug.Print "ViPickWindow::UpdateMe"

    If Not m_layeredAttributes Is Nothing Then
        m_layeredAttributes.Update Me.hWnd, m_layeredAttributes.theDC
    Else
        Me.Refresh
    End If
End Sub

Private Sub HandleCopyTab(ByVal newTabIndex As Long)
    If m_selectedItem Is Nothing Then
        Exit Sub
    End If

    If m_itemBank.GetCollectionByIndex(newTabIndex).Count > 1 Then
        m_itemBank.GetCollectionByIndex(newTabIndex).Add m_selectedItem.Clone, , 1
    Else
        m_itemBank.GetCollectionByIndex(newTabIndex).Add m_selectedItem.Clone
    End If
    
    RefreshRelevantPads m_itemBank.GetTabByIndex(newTabIndex)
End Sub

Private Sub HandleMoveTab(ByVal newTabIndex As Long)
    If m_selectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Not ExistInCol(m_currentItems, m_selectedItemIndex) Then
        MsgBox "ViPad caught a catastrophic error before it happend. Please close and re-open ViPad and then report error 1 to author", vbCritical
        Exit Sub
    End If
    
    m_currentItems.Remove m_selectedItemIndex
    
    If m_itemBank.GetCollectionByIndex(newTabIndex).Count > 1 Then
        m_itemBank.GetCollectionByIndex(newTabIndex).Add m_selectedItem, , 1
    Else
        m_itemBank.GetCollectionByIndex(newTabIndex).Add m_selectedItem
    End If
    
    PositionItems
    ReDraw
End Sub

Private Sub HandleContextMenuResult(contextMenuResult As Long)

    Select Case contextMenuResult
    
    Case IDMNU_DELETE
        If m_selectedItem Is Nothing Then
            Exit Sub
        End If
        
        'Test if it's in LNK bank or not
        If InStr(LCase(m_selectedItem.TargetPath), LCase(ApplicationDataPath)) = 1 Then
            DeleteFile m_selectedItem.TargetPath
        End If
        
        If FileExists(ApplicationDataPath & "\" & m_selectedItem.IconPath) Then
            m_selectedItem.ReleaseIcon
            
            'It doesn't matter if this function fails, really
            Call DeleteFile(ApplicationDataPath & "\" & m_selectedItem.IconPath)
        End If
            
        m_currentItems.Remove m_selectedItemIndex
        m_masterCollection.Remove m_selectedItem.GlobalIdentifer
        
        PositionItems
        ReDraw
        
        m_padChanged = True
        
    Case IDMNU_EDIT
        Dim itemInfo As New ChangeItemWindow
        itemInfo.Caption = m_selectedItem.Caption
        itemInfo.ItemCaption = m_selectedItem.Caption
        itemInfo.IconPath = ApplicationDataPath & "\" & m_selectedItem.IconPath
        itemInfo.Arguements = m_selectedItem.TargetArguements
        itemInfo.Target = m_selectedItem.TargetPath
        itemInfo.Show vbModal, Me
        
        If itemInfo.Changed Then
            m_selectedItem.Caption = itemInfo.ItemCaption
            Set m_selectedItem.Icon = itemInfo.IconImage
            m_selectedItem.IconPath = itemInfo.IconPath
            m_selectedItem.TargetArguements = itemInfo.Arguements
            m_selectedItem.TargetPath = itemInfo.Target
        
            Set itemInfo = Nothing
        End If
        
        PositionItems
        ReDraw
        
        m_padChanged = True
        
    Case IDMNU_QUIT
        EndApplication
        
    Case Else
        If contextMenuResult > 100 And contextMenuResult < 200 Then
            HandleMoveTab contextMenuResult - 100
        ElseIf contextMenuResult > 200 And contextMenuResult < 300 Then
            HandleCopyTab contextMenuResult - 200
        End If
    
    End Select

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
Dim cursorPos As win.POINTL
Dim menuResult As Long

    Debug.Print "Form_MouseUp: " & X & " -" & Y & " " & Button

    GetCursorPos cursorPos
    
    If m_mouseX = X And m_mouseY = Y Then
        If Not m_selectedItem Is Nothing Then
            If Button = vbLeftButton Then
                If Config.MinimizeAfterLauch Then
                    UpdateTabDimensionsIfPossible
                    ShowWindow Me.hWnd, SW_MINIMIZE
                End If
                
                Err.Clear
                Shell QuoteString(m_selectedItem.TargetPath) & " " & m_selectedItem.TargetArguements, vbNormalFocus
                If Err Then
                    Err.Clear
                    'MsgBox QuoteString("explorer.exe") & " " & QuoteString(m_selectedItem.TargetPath) & " " & m_selectedItem.TargetArguements
                    Shell QuoteString("explorer.exe") & " " & QuoteString(m_selectedItem.TargetPath) & " " & m_selectedItem.TargetArguements, vbNormalFocus
                End If
            ElseIf Button = vbRightButton Then
                m_contextMenuOpen = True
                
                UpdateTabMenu
                menuResult = m_itemOptionsMenu.ShowMenu(Me.hWnd, cursorPos.X, cursorPos.Y)
                'If menuResult = 0 Then m_contextMenuOpen = False
                
                If menuResult <> 0 Then
                    m_contextMenuOpen = False
                    HandleContextMenuResult menuResult
                    
                End If
            End If
        Else
            'm_trayIcon_onMouseUp CLng(Button)
            If Button = vbRightButton Then HandleTrayMenuContextMenuResult m_defaultMenu.ShowMenu(Me.hWnd, cursorPos.X, cursorPos.Y)
        End If
    Else
        If Button = vbLeftButton Then
            If Not m_dragItem Is Nothing Then
                If Y < m_headerOffset And m_showHeader Then
                    DropItemToPad (m_header.GetSelectedItemIndex + 1)
                Else
                    SwapItemOnPad GetSelectedItemIndex(X, Y)
                End If
                
                m_padChanged = True
            End If
        End If
    End If
End Sub

Private Sub SwapItemOnPad(proposedIndex As Long)


    If Not proposedIndex > -1 Then
        proposedIndex = m_oldSelectedItemIndex
    End If
    
    m_currentItems.Remove m_oldSelectedItemIndex
    
    If proposedIndex > m_currentItems.Count Then
        m_currentItems.Add m_dragItem
    Else
        m_currentItems.Add m_dragItem, , proposedIndex
    End If
    
    PositionItems
    ReDraw

    Set m_dragItem = Nothing

End Sub

Private Sub DropItemToPad(ByVal padIndex As Long)
    m_itemBank.GetCollectionByIndex(padIndex).Add m_dragItem
    m_currentItems.Remove m_oldSelectedItemIndex
    
    Set m_dragItem = Nothing
    m_dragMode = False
    
    PositionItems
    ReDraw
End Sub

Private Sub Form_Resize()

    If m_windowDimensions.cy = Me.ScaleHeight And _
        m_windowDimensions.cx = Me.ScaleWidth Then
        
        Exit Sub
    End If
    
    m_windowDimensions.cy = Me.ScaleHeight
    m_windowDimensions.cx = Me.ScaleWidth
    
    InitializeGrid
    Set m_rollOver = New GDIPGraphicPath
    
    m_header.WindowWidth = m_windowDimensions.cx
    
    m_pageSelector.X = (m_windowDimensions.cx / 2) - (m_pageSelector.Width / 2)
    m_pageSelector.Y = m_windowDimensions.cy - 20
        
    ReDraw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FileExists(m_tempLinkPath) Then
        DeleteFile m_tempLinkPath
        m_tempLinkPath = vbNullString
    End If
    
    If ExistInCol(ViPickWindows, m_systemId) Then ViPickWindows.Remove m_systemId
    If Not AppZorderKeeper Is Nothing Then AppZorderKeeper.RemoveChild Me
    
    Set m_header = Nothing

Dim thisRect As GdiPlus.RECTL

    If Not m_myTab Is Nothing Then
        If Not Config.InstanceMode Then
            Config.MainWindowRect = thisRect
            Config.WindowState = Me.WindowState
        End If
    End If
    
    Set m_masterCollection = Nothing
    
    If Not m_glassWindow Is Nothing Then
        Unload m_glassWindow
        Set m_glassWindow = Nothing
    End If
    
    If Not m_searchTextBox Is Nothing Then
        Unload m_searchTextBox
        Set m_searchTextBox = Nothing
    End If
    
    CheckViPadInstances
End Sub

Sub UpdateTabDimensionsIfPossible()
    If Me.WindowState <> vbNormal Then Exit Sub

Dim thisRect As GdiPlus.RECTL

    With thisRect
        .Height = Me.Height / Screen.TwipsPerPixelY
        .Width = Me.Width / Screen.TwipsPerPixelX
        .Left = Me.Left / Screen.TwipsPerPixelX
        .Top = Me.Top / Screen.TwipsPerPixelY
    End With

    If Not m_myTab Is Nothing Then
        m_myTab.Dimensions = Serialize_RectL(thisRect)
    End If
    
End Sub

Private Sub CheckViPadInstances()
    If AppRestarting Then Exit Sub

    If ViPickWindows.Count = 0 Then
        EndApplication
    End If
End Sub

Function IHookSink_WindowProc(hWnd As Long, msg As Long, wParam As Long, lParam As Long) As Long

On Error GoTo Handler

Dim bEatIt As Boolean
Dim sysCommand As Long
Dim pos As WINDOWPOS
Dim bEat As Boolean

    If msg = WM_EXITMENULOOP Then
        m_contextMenuOpen = True
        
    ElseIf msg = WM_KEYDOWN Then

    ElseIf msg = WM_SIZE Then
    
        If wParam = SIZE_MINIMIZED Then
            If Not m_glassWindow Is Nothing Then ShowWindow m_glassWindow.hWnd, SW_HIDE
        ElseIf wParam = SIZE_RESTORED Then
            If Not m_glassWindow Is Nothing Then ShowWindow m_glassWindow.hWnd, SW_SHOW
        End If
        
    ElseIf msg = WM_SYSCOMMAND Then
    
        sysCommand = (wParam And &HFFF0)

        If sysCommand = SC_SCREENSAVE Then
            DumpPads
        ElseIf sysCommand = SC_CLOSE Then
            Unload Me
        ElseIf sysCommand = SC_MINIMIZE Then
        
            If Not m_glassWindow Is Nothing Then ShowWindow m_glassWindow.hWnd, SW_HIDE
        ElseIf sysCommand = SC_RESTORE Then
            If Not m_glassWindow Is Nothing Then ShowWindow m_glassWindow.hWnd, SW_SHOW
        End If
        
    ElseIf msg = WM_ACTIVATE Then
    
        If Not wParam = WA_INACTIVE Then
            If m_glassMode Then
                If Not m_glassApplied Then
                    If ApplyGlassIfPossible(Me.hWnd) Then
                        m_glassApplied = True
                    End If
                End If
            End If
        
            If URLCatcher.Ready Then
                URLCatcher.Show
            End If
        Else
            Debug.Print "Reverting Search State!"
        
            If m_searchMode And Not m_previousItems Is Nothing Then
                m_searchMode = False
                Set m_searchResults = Nothing

                Set m_currentItems = m_previousItems
                PositionItems
                ReDraw
            End If
        End If

    ElseIf msg = WM_DWMCOMPOSITIONCHANGED Then
    
Dim dwmEnabled As Long
        
        IHookSink_WindowProc = 0
        bEat = True
        
        DwmIsCompositionEnabled dwmEnabled
        If m_glassMode Then
            UnApplyGlassIfPossible Me.hWnd
            m_glassApplied = ApplyGlassIfPossible(Me.hWnd)
        End If
        
    ElseIf msg = WM_DROPFILES Then

        Dim hFilesInfo As Long
        Dim szFileName As String
        Dim wTotalFiles As Long
        Dim wIndex As Long
        
        hFilesInfo = wParam
        wTotalFiles = DragQueryFileW(hFilesInfo, &HFFFF, ByVal 0&, 0)
    
        For wIndex = 0 To wTotalFiles
            szFileName = Space$(1024)
            
            If Not DragQueryFileW(hFilesInfo, wIndex, StrPtr(szFileName), Len(szFileName)) = 0 Then
                Form_DragDropFile TrimNull(szFileName)
            End If
        Next wIndex
        
        DragFinish hFilesInfo
        
    ElseIf msg = WM_WINDOWPOSCHANGING Then
    
        If Config.StickToDesktop Then
            Call win.CopyMemory(pos, ByVal lParam, Len(pos))
            
            'If IsHwndBelongToUs(pos.hwndInsertAfter) Then
            
            If Not pos.Flags And SWP_NOZORDER Then
                pos.hwndInsertAfter = HWND_BOTTOM
                'pos.Flags = pos.Flags Or SWP_NOZORDER
                 
                Call win.CopyMemory(ByVal lParam, pos, Len(pos))
            
            End If
            
            'End If
        End If
    
    End If
    
    If Not bEat Then
        ' Just allow default processing for everything else.
        IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wParam, lParam)
    End If

    Exit Function
Handler:
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
           InvokeWindowProc(hWnd, msg, wParam, lParam)
End Function

Private Function ChangesActiveIndex(newX As Long) As Boolean

Dim proposedPickIndex As Long
    
    proposedPickIndex = HowManyInto(newX * -1, m_windowDimensions.cx) + 1
    If proposedPickIndex <> m_activePickGridIndex Then
        ChangesActiveIndex = True
    End If

End Function

Function LeftButton() As Boolean
LeftButton = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Private Sub m_glassWindow_onDblClick()
    Form_DblClick
End Sub

Private Sub m_glassWindow_onKeyPress(KeyAscii As Long)
    Form_KeyPress CInt(KeyAscii)
End Sub

Private Sub m_glassWindow_onMouseDown(Button As Integer, X As Long, Y As Long)
    Form_MouseDown Button, 0, CSng(X), CSng(Y)
End Sub

Private Sub m_glassWindow_onMouseMove(Button As Integer, X As Long, Y As Long)
    Form_MouseMove Button, 0, CSng(X), CSng(Y)
End Sub

Private Sub m_glassWindow_onMouseUp(Button As Integer, X As Long, Y As Long)
    Form_MouseUp Button, 0, CSng(X), CSng(Y)
End Sub

Private Sub m_glassWindow_QueryOnItem(X As Long, Y As Long, onItem As Boolean)
    If Not m_selectedItem Is Nothing Or Not m_dragItem Is Nothing Then
        'prevents glass window from being able to move the window whilst attempting to drag
        
        onItem = True
        Exit Sub
    End If
    
    onItem = IsMouseCursorInsideDragableArea(X, Y)
End Sub

Private Sub m_header_onClick(targetItem As HeaderItem)
    If Not Config.InstanceMode Then
        ShowTabByIndex targetItem.ItemIndex + 1
    Else
        'If we are in instance mode, ignore index and select the tab assigned to this instance
        If Not m_myTab Is Nothing Then DisplayTab m_myTab
    End If
End Sub

Private Sub m_header_onDeleteHeader(ByVal headerItemIndex As Long)

    If Config.InstanceMode Then
        m_itemBank.RemoveTab Me.UniqueID
        Unload Me
    Else
        m_itemBank.RemoveIndex headerItemIndex + 1
        m_header.PopulateHeader m_itemBank
    End If
    
    If m_itemBank.GetCollectionCount > 0 Then Set m_myTab = m_itemBank.GetTabByIndex(1)
End Sub

Private Sub m_header_onNewHeader()
    CreateNewTab
End Sub

Private Sub m_header_onRenamedHeader(ByVal headerItemIndex As Long, ByVal newText As String)
    If Not Config.InstanceMode Then
        m_itemBank.SetCollectionAlias headerItemIndex + 1, newText
    Else
        m_myTab.Alias = newText
    End If
End Sub

Private Sub m_header_onSearchMode()
    SummonLocalSearchBox
End Sub

Private Sub m_header_onSwitchMode()
    ToggleAppType
End Sub

Private Sub m_header_RequestMeasureString(theText As String, theFont As GDIPFont, theWidth As Single)
    theWidth = m_graphics.MeasureString(theText, theFont).Width
End Sub

Private Sub m_header_RequestReDraw()
    If Not m_dragMode Then ReDraw
End Sub

Private Sub m_optionsWindow_onClose()
    m_optionsWindow.Hide
End Sub

Private Sub m_optionsWindow_onLayeredModeChanged()
    
Dim doLayeredMode As Boolean
    
    If Config.ForceLayeredMode Then
        doLayeredMode = True
    End If
    
    If doLayeredMode And m_glassWindow Is Nothing Then
    
        If m_glassApplied Then
            If UnApplyGlassIfPossible(Me.hWnd) Then
                m_glassApplied = False
            End If
        End If
        
        m_glassMode = False
        
        Set m_glassWindow = New GlassContainer
        m_glassWindow.DragDrop = True
        
        Set m_layeredAttributes = MakeLayerdWindow(Me, False)
        
        m_glassWindow.Show
        m_glassWindow.AttachForm Me

        HandleWindowStyle
        
        Me.Refresh

    ElseIf Not doLayeredMode Then
 
        If Not m_glassApplied And IsGlassAvailable Then

            m_glassMode = True
            
            m_glassWindow.ReleaseClient
            Unload m_glassWindow
            
            Set m_glassWindow = Nothing
            'Set m_layeredAttributes = Nothing
            
            m_layeredAttributes.SelectVBBitmap
            UnMakeLayeredWindow Me, m_layeredAttributes
            
            Set m_layeredAttributes = Nothing
            
            HandleWindowStyle
            UpdateMe
            
            Me.Hide
            Me.Show
            Me.Refresh

            If ApplyGlassIfPossible(Me.hWnd) Then
                m_glassApplied = True
            End If
        End If
        
    End If
End Sub

Private Sub m_optionsWindow_onNewIconSize(newSize As Long)
    m_IconSize = newSize
    
    CalculateCapacities
    PositionItems
    ReDraw
End Sub

Private Sub m_optionsWindow_onStickToDesktop()
    If Config.StickToDesktop Then
        SetWindowStyle Me.hWnd, False, WS_MINIMIZEBOX, False, False
        HideInTaskBar
    
        If Not m_glassWindow Is Nothing Then
            AppZorderKeeper.AddChild m_glassWindow
        Else
            AppZorderKeeper.AddChild Me
        End If
    Else
        SetWindowStyle Me.hWnd, False, WS_MINIMIZEBOX, True, False
        MakeVisibleInTaskBar
    End If
End Sub

Private Sub m_optionsWindow_onWindowStyleChanged()
    If Config.ShowControlBox Then
        SetWindowStyle Me.hWnd, False, WS_SYSMENU, True, False
    Else
        SetWindowStyle Me.hWnd, False, WS_SYSMENU, False, False
    End If
End Sub

Private Sub m_pageSelector_onSelectedItem(pageIndex As Long)
    
Dim pageDifference As Long

    If m_activePickGridIndex > pageIndex Then
        pageDifference = (m_activePickGridIndex - (pageIndex + 1))
    Else
        pageDifference = ((pageIndex + 1) - m_activePickGridIndex)
    End If

    If pageDifference = 0 Then
        pageDifference = 1
    End If
    
    m_mouseXOffset = 0
    m_pivotToIndex = pageIndex
    
    m_finalX = (m_windowDimensions.cx * m_pivotToIndex)
    m_xSlideSpeed = 50 * pageDifference
    
    timSlideAnimation.Enabled = True
    
End Sub

Private Sub m_searchTextBox_onChanged()
    If m_searchMode = False Then Exit Sub
    If m_searchTextBox Is Nothing Then Exit Sub
    If m_searchTextBox.Text = vbNullString Then Exit Sub
    
    SetSearchMode QueryByString(m_masterCollection, m_searchTextBox.Text)
        
    PositionItems
    ReDraw
End Sub

Private Sub m_searchTextBox_onClose()
    m_searchTextBox.Hide
    Me.SetFocus
End Sub

Private Sub m_trayIcon_onAddItem(newItemPath As String)
    Form_DragDropFile newItemPath
End Sub

Private Sub m_trayIcon_onMouseUp(Button As Long)

Dim menuResult As Long
Dim cursorPos As win.POINTL

    GetCursorPos cursorPos

    If Button = vbLeftButton Then
        'Me.WindowState = vbNormal
        SetForegroundWindow Me.hWnd
        
        'Me.Show
        If Me.WindowState = vbMinimized Then
            ShowWindow Me.hWnd, SW_RESTORE
        Else
            ShowWindow Me.hWnd, SW_SHOW
        End If
        
    ElseIf Button = vbRightButton Then
        SetForegroundWindow Me.hWnd
        
        menuResult = m_trayMenu.ShowMenu(Me.hWnd, cursorPos.X, cursorPos.Y)
        HandleTrayMenuContextMenuResult menuResult
    End If

End Sub

Private Sub HandleTrayMenuContextMenuResult(menuResult As Long)
    
    If menuResult = IDMNU_EXIT Then
        EndApplication
    ElseIf menuResult = IDMNU_RESTORE Then
        If Not m_glassWindow Is Nothing Then m_glassWindow.WindowState = vbNormal
        SetForegroundWindow Me.hWnd
        Me.Show
    ElseIf menuResult = IDMNU_SETTINGS Then
        m_optionsWindow.Show , Me
        
    ElseIf menuResult = IDMNU_ABOUT Then
        
        Dim A As New XMLWindow
        A.GoToURL "res://default_about"
        
        'If A.GoToURL("http://lee-soft.com/vipad/about_screens/" & GetVersionNumber & ".xml") = False Then
            
        'End If
        A.Show
    End If
    
End Sub

Private Sub MakeInternetShortcut(ByVal szCaption As String, ByVal szURL As String)

Dim thisPadItem As New LaunchPadItem
Dim theIcon As Long
Dim thisAlphaIcon As New AlphaIcon
Dim imageFileNamePath As String

    theIcon = IconHelper.GetIconFromFile(m_defaultBrowserIcon, SHIL_JUMBO)
    If IconHelper.IconIs48(m_defaultBrowserIcon) Then
        theIcon = IconHelper.GetIconFromFile(m_defaultBrowserIcon, SHIL_EXTRALARGE)
    End If
    
    thisAlphaIcon.CreateFromHICON theIcon
    imageFileNamePath = GenerateAvailableFileName

    If imageFileNamePath <> vbNullString Then
        thisAlphaIcon.Image.Save imageFileNamePath, GetPngCodecCLSID()
        thisPadItem.IconPath = GetFileNameFromPath(imageFileNamePath)
    End If
        
    thisPadItem.Caption = szCaption
    thisPadItem.TargetPath = m_defaultBrowserIcon
    thisPadItem.TargetArguements = szURL
    
    Set thisPadItem.Icon = thisAlphaIcon.Image
    'Keep Icon in Memory
    Set thisPadItem.AlphaIconStore = thisAlphaIcon
    
    AddLaunchItem thisPadItem
    
    PositionItems
    
    DrawItems
    
    RefreshHDC
    
    Unload m_urlCatcher
    
    Set m_selectedItem = thisPadItem
    HandleContextMenuResult IDMNU_EDIT
    
    Unload m_urlCatcher
    Set m_urlCatcher = Nothing

End Sub

Private Sub m_zOrderKeeper_onShowWindow()

    If Me.WindowState = vbMinimized Then
        If Not m_glassWindow Is Nothing Then ShowWindow m_glassWindow.hWnd, SW_RESTORE
        ShowWindow Me.hWnd, SW_RESTORE
    End If

    'StayOnTop Me, True
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    'TopMostViPadWindows

    If Not Config.TopMostWindow Then
        timTurnOfKeepOnTop.Enabled = True
    End If
    
End Sub

Private Sub Timer1_Timer()
    m_searchTextBox_onChanged
End Sub

Private Sub timActivateDelay_Timer()

End Sub

Private Sub timHoldCount_Timer()

    If Not LeftButton Then
        m_holdTicker = 0
        
        timHoldCount.Enabled = False
        Exit Sub
    End If
    
    If m_holdTicker < 4 Then
        m_holdTicker = m_holdTicker + 1
        
        Exit Sub
    End If
    
    timHoldCount.Enabled = False
    PullOutSelected
End Sub

Private Sub CreateTempLink()
    If FileExists(m_tempLinkPath) Then
        DeleteFile m_tempLinkPath
        m_tempLinkPath = vbNullString
    End If

    If m_selectedItem Is Nothing Then Exit Sub

Dim tempLink As New ShellLinkClass

    m_tempLinkPath = App.Path & "\" & m_selectedItem.Caption & ".lnk"
    Debug.Print m_tempLinkPath
    
    tempLink.CreateShortcut App.Path & "\" & m_selectedItem.Caption & ".lnk"
    tempLink.Target = m_selectedItem.TargetPath
    tempLink.Arguments = m_selectedItem.TargetArguements
    tempLink.Save
End Sub

Private Sub PullOutSelected()
    On Error Resume Next

Dim proposedIndex As Long

    If m_selectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Config.InstanceMode Then
    
        CreateTempLink
        Me.OLEDrag
        
        m_currentItems.Remove m_selectedItemIndex
    
        PositionItems
        ReDraw
        Set m_selectedItem = Nothing
        
    Else
    
        Set m_dragItem = m_selectedItem
        m_currentItems.Remove m_selectedItemIndex
    
        If m_selectedItemIndex = (m_currentItems.Count + 1) Then
            m_currentItems.Add m_blankItem
            m_oldSelectedItemIndex = m_currentItems.Count
        Else
            m_currentItems.Add m_blankItem, , m_selectedItemIndex
            m_oldSelectedItemIndex = m_selectedItemIndex
            
            Debug.Print "Caption:: " & m_currentItems(m_oldSelectedItemIndex).Caption
        End If
        
        'm_currentItems.Add m_dragItem
        
        PositionItems
        ReDraw
        
        m_dragOffsetX = m_dragItem.X - m_mouseX + m_gridOffsetX
        m_dragOffsetY = m_dragItem.Y - m_mouseY + m_headerOffset
        
        HandleDragItem m_mouseX, m_mouseY, GetSelectedItemIndex(m_mouseX, m_mouseY)
        Set m_selectedItem = Nothing
    
    End If
End Sub

Private Sub timSlideAnimation_Timer()

Dim finalXPosition As Long

    'm_pivotToIndex = CLng((m_currentBitmap_XOffset - (X_MARGIN)) / m_windowDimensions.cx) * -1
    finalXPosition = m_finalX
    
    'Debug.Print "finalXPosition:: " & finalXPosition
    
    'Debug.Print "slideAni: " & (m_pivotToIndex * m_windowDimensions.cx) & " - " & (m_currentBitmap_XOffset * -1)
    If m_currentBitmap_XOffset <> finalXPosition And m_currentBitmap_XOffset <> -finalXPosition Then
    
        'Debug.Print m_pivotToIndex & ":" & m_activePickGridIndex
    
        If m_pivotToIndex < m_activePickGridIndex Then

            If finalXPosition > (m_currentBitmap_XOffset + m_xSlideSpeed) Then

                If Not CBool((-finalXPosition - (m_currentBitmap_XOffset)) > m_xSlideSpeed) And _
                    ChangesActiveIndex(m_currentBitmap_XOffset + m_xSlideSpeed) Then

                    CurrentBitmapX = -finalXPosition
                Else
                    CurrentBitmapX = m_currentBitmap_XOffset + m_xSlideSpeed
                End If
            Else
                CurrentBitmapX = finalXPosition
            End If
            
        ElseIf m_pivotToIndex >= m_activePickGridIndex Then
        
            If finalXPosition > ((m_currentBitmap_XOffset - m_xSlideSpeed) * -1) Then
                CurrentBitmapX = m_currentBitmap_XOffset - m_xSlideSpeed
            Else
                CurrentBitmapX = -finalXPosition
            End If
        End If
    End If
    
End Sub

Sub AddShortcutsOnDesktop()

'Dim FSO As New FileSystemObject
Dim theDesktop As Scripting.Folder
Dim thisFile As Scripting.File
Dim theFiles As Collection

    Set theFiles = New Collection

    Set theDesktop = FSO.GetFolder(GetUserDesktopPath())
    For Each thisFile In theDesktop.Files
        If Right(UCase(thisFile.Name), 4) = ".LNK" Then
            theFiles.Add thisFile.Path, thisFile.Path
        End If
    Next
    
    Set theDesktop = FSO.GetFolder(GetPublicDesktopPath())
    For Each thisFile In theDesktop.Files
        If Right(UCase(thisFile.Name), 4) = ".LNK" Then
            If Not ExistInCol(theFiles, thisFile.Path) Then
                theFiles.Add thisFile.Path, thisFile.Path
            End If
        End If
    Next
    
    While theFiles.Count > 1
        DoEvents
        
        Form_DragDropFile CStr(theFiles(1))
        theFiles.Remove 1
    Wend
    

End Sub

Private Sub timTurnOfKeepOnTop_Timer()
Dim cursorPos As win.POINTL

    'StayOnTop Me, False
    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)

    SetForegroundWindow Me.hWnd
    SetActiveWindow Me.hWnd
    SendKeys "{TAB}"
    
    timTurnOfKeepOnTop.Enabled = False
End Sub

Private Sub SummonLocalSearchBox(Optional ByVal szInitialText As String)
    If m_searchTextBox Is Nothing Then Exit Sub

    m_searchTextBox.Text = szInitialText
    m_searchTextBox.Activate
    
    If m_glassWindow Is Nothing Then
        If Me.ScaleWidth > 200 Then
            MoveWindow m_searchTextBox.hWnd, _
                        (Me.Left / Screen.TwipsPerPixelX) + 60, _
                        (Me.Top / Screen.TwipsPerPixelY) + 30, _
                        Me.ScaleWidth - 104, 55, APIFALSE
        Else
            MoveWindow m_searchTextBox.hWnd, (Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY) + 30, Me.ScaleWidth, 55, APIFALSE
        End If
        
        m_searchTextBox.Show vbModeless, Me
    Else
        If Me.ScaleWidth > 200 Then
            MoveWindow m_searchTextBox.hWnd, (Me.Left / Screen.TwipsPerPixelX) + 60, (Me.Top / Screen.TwipsPerPixelY) + 10, Me.ScaleWidth - 120, 55, APIFALSE
        Else
            MoveWindow m_searchTextBox.hWnd, (Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY) + 10, Me.ScaleWidth, 55, APIFALSE
        End If
        
        m_searchTextBox.Show vbModeless, Me
    End If
    
    m_searchTextBox.Show vbModeless, Me

End Sub

Private Sub SetSearchMode(ByRef searchResults As Collection)
    If Not m_searchMode Then
        Set m_previousItems = m_currentItems
        m_searchMode = True
    End If
    
    Set m_searchResults = searchResults
    Set m_currentItems = searchResults
End Sub

Sub RefreshContents()
    PositionItems
    ReDraw
End Sub

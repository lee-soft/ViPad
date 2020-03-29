VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViPickHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Current Mouse Co-ordinates (relative to form)
Private m_lastMouseDownX As Long
Private m_lastMouseDownY As Long
Private m_mouseXOffset As Long
Private m_headerPositionY As Single

Private Const HEADER_SEPARATOR As Long = 15
Private Const HEADER_TEXT_SIZE As Long = 20

Public thePath As GDIPGraphicPath
Public thePath2 As GDIPGraphicPath
Public rollOver As GDIPGraphicPath
Public DragMode As Boolean

Public VirginRollover As Boolean

Private m_contextMenu As ContextMenu

Private m_headerItems As Collection
Private m_headerItemWidth As Long

Private m_fontFamilyHeader As GDIPFontFamily
Private m_fontHeader As GDIPFont

Private m_windowWidth As Long
Private m_Left As Long

Private m_itemCount As Long

Private m_currentSelectedItem As HeaderItem
Private m_clickedItem As HeaderItem

Private m_headerCentreItem As RECTF
Private m_gdiFormatCentre As Long
Private m_headerCentreText As String
Private m_lastKnownIndex As Long
Private m_renameItem As HeaderItem

Private m_dragOffset As Long
Private m_mouseClicked As Boolean

Private WithEvents m_renameWindow As APITextBox
Attribute m_renameWindow.VB_VarHelpID = -1

Private m_hWnd As Long
Private m_parentForm As Form
Private m_addNewItem As HeaderItem

Private m_deviant As Boolean
Private m_orignalWidth As Long
Private m_deviated As Long
Private m_unclicked As Boolean

Public Event RequestMeasureString(theText As String, theFont As GDIPFont, ByRef theWidth As Single)
Public Event RequestReDraw()

Public Event onClick(ByRef targetItem As HeaderItem)
Public Event onRenamedHeader(ByVal headerItemIndex As Long, ByVal newText As String)
Public Event onDeleteHeader(ByVal headerItemIndex As Long)
Public Event onNewHeader()
Public Event onSwitchMode()
Public Event onSearchMode()

Public Function SetSingleInstanceMode(ByVal szInstanceTitle As String)
    AddItemHeader szInstanceTitle, False
    Seal
End Function

Public Property Let Y(newY As Single)
    m_headerPositionY = newY
    Me.WindowWidth = m_windowWidth
End Property

Public Function ResetHeader()
    m_unclicked = True
End Function

Public Property Get Height() As String
    Height = HEADER_TEXT_SIZE
End Property

Public Property Let ParentForm(newParent As Form)
    Set m_parentForm = newParent
    m_hWnd = newParent.hWnd
End Property

Public Function SetClickedItemByIndex(ByVal newClickedIndex As Long)

Dim thisItem As HeaderItem

    For Each thisItem In m_headerItems
        If thisItem.ItemIndex = newClickedIndex Then
            Set m_clickedItem = thisItem
            Exit For
        End If
    Next
    
    GeneratePaths
    RaiseEvent RequestReDraw
End Function

Public Function ResetRollover()
    If Not m_currentSelectedItem Is Nothing Then
        m_lastKnownIndex = -1
    End If
    
    Set m_currentSelectedItem = Nothing
    
    GeneratePaths
End Function

Private Function GetSelectedItem(ByVal X As Single) As HeaderItem

Dim thisItem As HeaderItem

Dim startX As Long
Dim endX As Long

    startX = m_Left

    For Each thisItem In m_headerItems
        endX = startX + thisItem.GetOrigin.Width + HEADER_SEPARATOR
        
        If X > startX And X < endX Then
            If thisItem.Text = "|" Then Exit Function
        
            Set GetSelectedItem = thisItem
            Exit For
        End If
        
        startX = endX
    Next

End Function

Private Function RebuildContextMenu(includeDelete As Boolean)

    m_contextMenu.Clear

    If includeDelete Then m_contextMenu.AddItem 1, "Rename", NORMALITEM
    If includeDelete Then m_contextMenu.AddItem 2, "Delete", NORMALITEM
    m_contextMenu.AddItem 3, "New Tab", NORMALITEM
    m_contextMenu.AddSeperater
    m_contextMenu.AddItem 6, "Search ViPad", NORMALITEM
    m_contextMenu.AddSeperater
    
    If Config.InstanceMode Then
        m_contextMenu.AddItem 5, "Switch to Tabbed Mode"
    Else
        m_contextMenu.AddItem 5, "Switch to Instance Mode"
    End If
    
    m_contextMenu.AddSeperater
    m_contextMenu.AddItem 4, "Reset", NORMALITEM

End Function

Public Function MouseDown(Button As Integer, X As Long, Y As Long)
Dim windowRect As RECT
Dim contextResult As Long

    m_lastMouseDownY = Y
    m_lastMouseDownX = X

    GetWindowRect m_hWnd, windowRect

    If Button = vbLeftButton Then
        m_mouseClicked = True
    
        If Not m_currentSelectedItem Is Nothing Then
            m_unclicked = False
        
            Set m_clickedItem = m_currentSelectedItem
            
            If m_currentSelectedItem Is m_addNewItem Then
                RaiseEvent onNewHeader
            Else
                RaiseEvent onClick(m_currentSelectedItem)
            End If
        End If
    ElseIf Button = vbRightButton Then
        If m_currentSelectedItem Is Nothing Then
            RebuildContextMenu False
        Else
            If m_currentSelectedItem.ItemIndex = m_itemCount Then
                RebuildContextMenu False
            Else
                RebuildContextMenu True
            End If
        End If
        
        contextResult = m_contextMenu.ShowMenu(m_hWnd)
    
        If contextResult = 1 Then
        
            'MsgBox m_currentSelectedItem.Text
            If Not m_renameWindow Is Nothing Then
                Unload m_renameWindow
                Set m_renameWindow = Nothing
            End If
        
            Set m_renameWindow = New APITextBox
            Set m_renameItem = m_currentSelectedItem
        
            MoveWindow m_renameWindow.hWnd, windowRect.Left + X, windowRect.Top + Y, 300, 50, True
            m_renameWindow.Text = m_currentSelectedItem.Text
            m_renameWindow.Show vbModeless, m_parentForm
            m_renameWindow.SelectAll
            
        ElseIf contextResult = 2 Then
        
            RaiseEvent onDeleteHeader(m_currentSelectedItem.ItemIndex)
        ElseIf contextResult = 3 Then
        
            RaiseEvent onNewHeader

        ElseIf contextResult = 4 Then
            
            m_dragOffset = 0
    
            SetLeft
            SetItemHeaderOrigins
            GeneratePaths
            
            RaiseEvent RequestReDraw
        ElseIf contextResult = 5 Then
            RaiseEvent onSwitchMode
        ElseIf contextResult = 6 Then
            RaiseEvent onSearchMode
        End If
    End If

End Function

Public Function MouseMove(Button As Integer, X As Long, Y As Long)

Dim selectedItem As HeaderItem

    If Button = vbLeftButton And m_mouseClicked Then
    
        If Not IsEqualWithinReason(m_lastMouseDownX, X, 10) Then
            If m_mouseXOffset = 0 Then
                m_mouseXOffset = X - (m_dragOffset - m_deviated)
                
                
                Debug.Print "m_mouseXOffset:: " & m_mouseXOffset
            End If
            
            m_dragOffset = X - m_mouseXOffset
            m_deviated = 0
            m_deviant = False
            
            SetLeft
            SetItemHeaderOrigins
            
            GeneratePaths
            RaiseEvent RequestReDraw
            
        End If
        
        Exit Function
    End If

    m_mouseXOffset = 0
    m_mouseClicked = False
    
    Set selectedItem = GetSelectedItem(X)
    
    If Not selectedItem Is m_currentSelectedItem Then
        Set m_currentSelectedItem = selectedItem
        
        Set rollOver = New GDIPGraphicPath
        
        If Not selectedItem Is Nothing Then
            rollOver.AddArc selectedItem.GetOrigin.Left, m_headerPositionY, 32, 22, 135, 90
            rollOver.AddArc selectedItem.GetOrigin.Left + (selectedItem.GetOrigin.Width - 28) - 4, m_headerPositionY, 32, 22, -45, 90
        End If
    
        If Not m_unclicked Then GeneratePaths
        'GeneratePaths
        RaiseEvent RequestReDraw
    End If

End Function

Public Property Get Width() As Long
    Width = m_headerItemWidth
End Property

Public Property Let WindowWidth(newWidth As Long)
    'Debug.Print GetPotentialLeft(newWidth)

Dim oldWidth As Long
    oldWidth = m_windowWidth

    m_windowWidth = newWidth
    
    If IsOffLeftEdge(GetPotentialLeft(newWidth)) = False Then
        m_deviant = False
        SetLeft
    Else
        If m_Left > 0 Then
            m_Left = 0
        Else
            If GetPotentialLeft(newWidth) > m_Left Then
                m_deviant = False

                SetLeft
            Else
                If m_deviant = False Then
                    m_orignalWidth = oldWidth
                    m_deviant = True
                End If
            
                'This works but why?
                m_deviated = -(m_orignalWidth - m_windowWidth) * 0.5
            End If
        End If
    End If
    
    SetItemHeaderOrigins
    GeneratePaths
End Property

Public Function GetSelectedItemIndex() As Long
    If m_currentSelectedItem Is Nothing Then
        GetSelectedItemIndex = -1
        Exit Function
    End If

    GetSelectedItemIndex = m_currentSelectedItem.ItemIndex
End Function

Private Function GeneratePaths()

Dim thisItem As HeaderItem
Dim thisOrigin As RECTF

    Set thePath = New GDIPGraphicPath
    Set thePath2 = New GDIPGraphicPath
    
    For Each thisItem In m_headerItems
        thisOrigin = thisItem.GetOrigin
    
        If m_clickedItem.ItemIndex = thisItem.ItemIndex Then
            If GetSelectedItemIndex = thisItem.ItemIndex Then
                thePath.AddString thisItem.DisplayedText, m_fontFamilyHeader, FontStyle.FontStyleBold, HEADER_TEXT_SIZE, thisOrigin, 0
            End If
        Else
            thePath.AddString thisItem.DisplayedText, m_fontFamilyHeader, FontStyle.FontStyleBold, HEADER_TEXT_SIZE, thisOrigin, 0
        End If
        
        thisOrigin.Left = thisOrigin.Left + 1.5
        thisOrigin.Top = thisOrigin.Top + 1.5
        
        'If Not m_clickedItem.ItemIndex = thisItem.ItemIndex Then
        thePath2.AddString thisItem.DisplayedText, m_fontFamilyHeader, FontStyle.FontStyleBold, HEADER_TEXT_SIZE, thisOrigin, 0
    Next

End Function

Private Function AddItemHeader(ByVal theText As String, ByVal theLastItem As Boolean) As HeaderItem

Dim thisItem As HeaderItem
Dim itemWidth As Single


    Set thisItem = New HeaderItem
    Set AddItemHeader = thisItem
    
    thisItem.Text = theText
    
    RaiseEvent RequestMeasureString(theText, m_fontHeader, itemWidth)
    thisItem.SetOrigin 0, 0, 0, itemWidth

    If theText = "|" Then
        thisItem.ItemIndex = -1
    Else
        thisItem.ItemIndex = m_itemCount
    End If

    m_headerItems.Add thisItem
    m_headerItemWidth = m_headerItemWidth + thisItem.GetOrigin.Width + HEADER_SEPARATOR
    
    If Not theLastItem And theText <> "|" Then
        m_itemCount = m_itemCount + 1
        AddItemHeader "|", False
    End If

End Function

Private Function CutRightUntilFits(targetLeft As Long, ByVal sourceString As String) As String
    
Dim thisItemWidth As Long
Dim originalWidth As Single
Dim requestedWidth As Long

Dim newString As String
Dim newWidth As Single
Dim cutAmount As Long

    'Debug.Print "CutRightUntilFits:: " & sourceString

    RaiseEvent RequestMeasureString(sourceString, m_fontHeader, originalWidth)
    requestedWidth = originalWidth + (m_windowWidth - targetLeft) - 10
    
    newString = sourceString
    
    If requestedWidth > 23 Then
    
        cutAmount = 0
        
        Do
            cutAmount = cutAmount + 1
            newString = Mid(sourceString, 1, Len(sourceString) - cutAmount)
            RaiseEvent RequestMeasureString(newString, m_fontHeader, newWidth)
            
            'MsgBox newWidth & ":" & requestedWidth
            
        Loop While (newWidth > requestedWidth)
    Else
        newString = ""
    End If
    
    CutRightUntilFits = newString & "..."
    
    
    
End Function

Public Function SetLeft()
    m_Left = m_dragOffset + (m_windowWidth / 2) - (m_headerItemWidth / 2) + (HEADER_SEPARATOR / 2)
End Function

Public Function GetPotentialLeft(newWidth As Long)
    GetPotentialLeft = m_dragOffset + (newWidth / 2) - (m_headerItemWidth / 2) + (HEADER_SEPARATOR / 2)
End Function

Private Function IsOffLeftEdge(newLeft As Long) As Boolean

Dim thisItem As HeaderItem
Dim thePath As New GDIPGraphicPath
Dim theXOffset As Single
Dim extraXOffset As Long

    theXOffset = newLeft
    IsOffLeftEdge = False
    
    m_headerCentreItem.Width = m_windowWidth
    m_headerCentreItem.Height = 100

    For Each thisItem In m_headerItems
        If theXOffset < 1 Then
            'Stop moving header
            'set dragoffset in a such a way that its never off the screen
            
            
            IsOffLeftEdge = True
        End If
        
        theXOffset = theXOffset + thisItem.GetOrigin.Width + HEADER_SEPARATOR
    Next

End Function

Private Function SetItemHeaderOrigins() As Boolean

Dim thisItem As HeaderItem
Dim thePath As New GDIPGraphicPath
Dim theXOffset As Single
Dim extraXOffset As Long

    theXOffset = m_Left
    
    m_headerCentreItem.Width = m_windowWidth
    m_headerCentreItem.Height = 100

    For Each thisItem In m_headerItems
        If theXOffset < 0 Then
            'Stop moving header
            'set dragoffset in a such a way that its never off the screen
            
            thisItem.DisplayedText = thisItem.Text
            thisItem.SetOrigin 20, theXOffset, m_headerPositionY, thisItem.GetOrigin.Width
            
        ElseIf (theXOffset + thisItem.GetOrigin.Width) > m_windowWidth Then
            
            If thisItem.Text <> "|" Then
                thisItem.DisplayedText = CutRightUntilFits(CLng((theXOffset + thisItem.GetOrigin.Width)), thisItem.Text)
            Else
                thisItem.DisplayedText = thisItem.Text
                thisItem.DisplayedText = ""
            End If
            
            thisItem.SetOrigin 20, theXOffset, m_headerPositionY, thisItem.GetOrigin.Width
        Else
            thisItem.DisplayedText = thisItem.Text
            thisItem.SetOrigin 20, theXOffset, m_headerPositionY, thisItem.GetOrigin.Width
        End If
        
        theXOffset = theXOffset + thisItem.GetOrigin.Width + HEADER_SEPARATOR
    Next

End Function

Public Function ReCalculateWidth()
    
Dim thisItem As HeaderItem
Dim itemWidth As Single

    m_headerItemWidth = 0

    For Each thisItem In m_headerItems
        RaiseEvent RequestMeasureString(thisItem.Text, m_fontHeader, itemWidth)
        
        thisItem.SetOrigin 0, 0, 0, itemWidth
        m_headerItemWidth = m_headerItemWidth + thisItem.GetOrigin.Width + HEADER_SEPARATOR
    Next

End Function

Public Function PopulateHeader(ByRef srcBank As ViBank)
    
    Set m_headerItems = New Collection
    m_itemCount = 0
    m_headerItemWidth = 0

Dim thisCollectionAlias As String
Dim thisCollectionIndex As Long
Dim thisTab As ViTab

    For thisCollectionIndex = 1 To srcBank.GetCollectionCount
        Set thisTab = srcBank.GetTabByIndex(thisCollectionIndex)

        AddItemHeader thisTab.Alias, False
    Next
    
    Seal
    
    'Does all the redrawing here
    Me.WindowWidth = m_windowWidth
End Function

Private Function Seal()
    Set m_addNewItem = AddItemHeader("+", True)
End Function

Private Sub Class_Initialize()
    Set m_fontHeader = New GDIPFont
    Set m_fontFamilyHeader = New GDIPFontFamily
    Set m_headerItems = New Collection
    
    Set rollOver = New GDIPGraphicPath
    Set m_clickedItem = New HeaderItem
    Set m_contextMenu = New ContextMenu
    Set m_renameWindow = New APITextBox
    
    RebuildContextMenu False

    m_fontHeader.Constructor m_fontFamilyHeader, HEADER_TEXT_SIZE, FontStyleBold
    m_fontFamilyHeader.Constructor "Arial"
    
    GdipCreateStringFormat 0, 0, m_gdiFormatCentre
    GdipSetStringFormatAlign m_gdiFormatCentre, StringAlignmentCenter
    
    GeneratePaths
    
End Sub

Private Sub Class_Terminate()
    If Not m_renameWindow Is Nothing Then
        Unload m_renameWindow
        Set m_renameWindow = Nothing
    End If
    
    If m_gdiFormatCentre <> 0 Then GdipDeleteStringFormat m_gdiFormatCentre
End Sub

Private Sub DeleteCurrentItem()


End Sub

Private Sub m_renameWindow_onClose()

    m_renameWindow.Hide

    If Not m_renameItem Is Nothing Then
        m_renameItem.Text = m_renameWindow.Text
        RaiseEvent onRenamedHeader(m_renameItem.ItemIndex, m_renameItem.Text)
        
        ReCalculateWidth
        ResetRollover
        
        SetItemHeaderOrigins
        GeneratePaths

        RaiseEvent RequestReDraw
    End If
End Sub

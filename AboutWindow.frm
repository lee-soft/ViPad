VERSION 5.00
Begin VB.Form XMLWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Untitled Window"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2700
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   135
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timers 
      Enabled         =   0   'False
      Index           =   0
      Left            =   1740
      Top             =   1320
   End
   Begin ViPad.GraphicImage Images 
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   -2220
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2143
   End
   Begin ViPad.GraphicLabel Labels 
      Height          =   345
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   -1500
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ViPad.GraphicButton Buttons 
      Height          =   450
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   -3000
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "XMLWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Option Explicit

Private Const AlignDefault As String = "centre"

Private m_appIcon As AlphaIcon
Private m_creditsXML As DOMDocument

Private m_Y As Long
Private m_X As Long
Private m_glassMode As Boolean

Private m_theWindow As XMLWindowType

Public returnCode As Long

Private Sub ParseTimer(ByRef theTimer As IXMLDOMElement)

    Load Timers(Timers.count)
    With Timers(Timers.UBound)
        .Interval = CLng(theTimer.getAttribute("interval"))
        .Tag = theTimer.getAttribute("command")
        .Enabled = True
    End With

End Sub

Private Sub ParseH(ByRef theH1 As IXMLDOMElement, theSize As Long)

Dim theCaption As String
Dim theAlign As String
Dim theFontStyle As Long
Dim realPosition_X As Long

Dim theLabelIndex As Long

    theCaption = theH1.Text
    theFontStyle = fontStyle.FontStyleBold
    
    If Not IsNull(theH1.getAttribute("align")) Then
        theAlign = theH1.getAttribute("align")
    End If
    
    theLabelIndex = AddText(theCaption, theAlign, theFontStyle, 10 + (theSize * 5))
    m_Y = m_Y + Labels(theLabelIndex).Height

End Sub

Private Sub ParseParagraph(ByRef theText As IXMLDOMElement)

Dim theCaption As String
Dim theAlign As String
Dim theFontStyle As Long
Dim realPosition_X As Long

Dim theDimensions As RECTF

Dim theLabelIndex As Long

    If Me.FontBold Then
        theFontStyle = fontStyle.FontStyleBold
    ElseIf Me.FontItalic Then
        theFontStyle = fontStyle.FontStyleItalic
    Else
        theFontStyle = fontStyle.FontStyleRegular
    End If

    theCaption = theText.Text
    
    If Not IsNull(theText.getAttribute("align")) Then
        theAlign = theText.getAttribute("align")
    End If
    
    theLabelIndex = AddText(theCaption, theAlign, theFontStyle, Me.Font.Size)
    m_Y = m_Y + Labels(theLabelIndex).Height
    

End Sub

Private Function AddText(ByVal theCaption As String, ByVal theAlign As String, ByVal theFontStyle As Long, ByVal theFontSize As Long) As Long
    If theCaption = vbNullString Then Exit Function

    Load Labels(Labels.count)

Dim realPosition_X As Long


    If theAlign = vbNullString Then
        theAlign = AlignDefault
    End If
    
    With Labels(Labels.UBound)
    
        .BackColor = Me.BackColor
    
        If m_glassMode Then
            .ForeColor = vbWhite
        End If
            
        .Font.Size = theFontSize
        .Caption = Replace(theCaption, "%BUILDINFO%", App.Major & "." & App.Minor & "." & App.Revision)
        .AutoSize = True
        
        If .Width > Me.ScaleWidth Then
            .AutoSize = False
            .MultiLine = True
            .Width = Me.ScaleWidth
            .AutoSize = True

        End If
        
    If LCase(theAlign) = "centre" Then
        realPosition_X = (Me.ScaleWidth / 2) - (.Width / 2)
    ElseIf LCase(theAlign) = "right" Then
        realPosition_X = (Me.ScaleWidth) - .Width
    End If
        
        
        .Top = m_Y
        .Left = realPosition_X
        
        .Visible = True
    End With
    
    'm_rectPosition.Top = m_Y
    'm_rectPosition.Left = realPosition_X
    'm_rectPosition.Width = m_rectPosition.Left + theDimensions.Width
    'm_rectPosition.Height = m_rectPosition.Top + theDimensions.Height
    'm_path.AddString theCaption, m_fontF, theFontStyle, CSng(theFontSize), m_rectPosition, 0
    
    AddText = Labels.UBound

End Function

Private Sub ParseButton(ByRef theButton As IXMLDOMElement)

Dim theCaption As String
Dim theCommand As String
Dim X As Long
Dim Y As Long
Dim theWidth As Long
Dim theHeight As Long

    theCaption = theButton.getAttribute("caption")
    X = theButton.getAttribute("x")
    Y = theButton.getAttribute("y")
    theWidth = theButton.getAttribute("width")
    theHeight = theButton.getAttribute("height")
    theCommand = theButton.getAttribute("command")

    Load Buttons(Buttons.count)
    With Buttons(Buttons.UBound)
        .Top = Y
        .Left = X
        .Width = theWidth
        .Height = theHeight
        .Caption = theCaption
        .Visible = True
        .Command = theCommand
    End With

End Sub

Private Sub ParseImage(ByRef theImage As IXMLDOMElement)

Dim thisImage As New GDIPImage
Dim theSrc As String
Dim theAlign As String
Dim realPosition_X As Long
Dim floatsY As Boolean

    Load Images(Images.count)
    With Images(Images.UBound)
    
    theSrc = theImage.getAttribute("src")
    
    If Not IsNull(theImage.getAttribute("align")) Then
        theAlign = theImage.getAttribute("align")
    End If
    
    If Not IsNull(theImage.getAttribute("float")) Then
        If LCase(theImage.getAttribute("float")) = "y" Then
            floatsY = True
        End If
    End If
    
    If theSrc = "APPICON" Then
        Set m_appIcon = New AlphaIcon
        m_appIcon.CreateFromHICON IconHelper.GetIconFromFile(App.Path & "\" & App.EXEName & ".exe", SHIL_JUMBO)
        Set thisImage = m_appIcon.Image
    End If
    
    If Not thisImage Is Nothing Then
        If LCase(theAlign) = "centre" Then
            realPosition_X = (Me.ScaleWidth / 2) - (thisImage.Width / 2)
        ElseIf LCase(theAlign) = "right" Then
            realPosition_X = (Me.ScaleWidth) - thisImage.Width
        End If
        
        .Image = thisImage
        .Left = realPosition_X
        .Top = m_Y
        .Visible = True
        
        'm_graphicsImage.DrawImage thisImage, realPosition_X, m_Y, thisImage.width, thisImage.height
        If Not floatsY Then m_Y = m_Y + thisImage.Height
    End If

    End With
    
End Sub

Private Sub ParseCreditsXML()
    'On Error GoTo Handler
    
Dim thisElement As IXMLDOMElement
Dim theWidth As Long
Dim theHeight As Long
Dim windowTitle As String
Dim theWindowXML As IXMLDOMElement

    Set theWindowXML = m_creditsXML.selectSingleNode("xmlwindow")
    m_theWindow = ParseXMLWindow(theWindowXML)

    theWidth = m_theWindow.theWidth
    theHeight = m_theWindow.theHeight
    windowTitle = m_theWindow.theTitle

    Me.Caption = windowTitle

    Me.Width = theWidth * Screen.TwipsPerPixelX
    Me.Height = theHeight * Screen.TwipsPerPixelY
    
    'm_graphicsImage.SmoothingMode = SmoothingModeHighQuality
    'm_graphicsImage.InterpolationMode = InterpolationModeHighQualityBicubic

    For Each thisElement In theWindowXML.childNodes
        
        Select Case LCase(thisElement.tagName)
        
        Case "img"
            ParseImage thisElement
            
        Case "p"
            ParseParagraph thisElement
            
        Case "h1"
            ParseH thisElement, 1
        Case "h2"
            ParseH thisElement, 2
        Case "h3"
            ParseH thisElement, 3

        Case "button"
            ParseButton thisElement
            
        Case "timer"
            ParseTimer thisElement
        
        End Select
    Next

    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical, "Error parsing XMLWindow"

End Sub

Private Sub Buttons_onClick(index As Integer)

Dim theCommand As CommandType

    theCommand = ParseCommandText(Buttons(index).Command)
    ExecuteCommand Me, theCommand.commandText, theCommand.Args
End Sub

Private Sub Form_Click()

    If m_theWindow.onClick.commandText <> "" Then
        ExecuteCommand Me, m_theWindow.onClick.commandText, m_theWindow.onClick.Args
    End If
End Sub

Private Sub Form_DblClick()
'    Unload Me
End Sub

Private Sub Form_Initialize()
    Set m_appIcon = New AlphaIcon
    Set m_creditsXML = New DOMDocument
    
    'm_creditsXML.XML = HttpRequest("http://lee-soft.com/vipick/versions/1.xml")
    'm_creditsXML.XML = Fload("C:\test.xml")

    'ParseCreditsXML
    
    'GoToURL "res://ABOUT"
End Sub

Private Function ResetWindow()

    Set m_creditsXML = New DOMDocument
    m_Y = 0
    m_X = 0

    While Timers.count > 1
        Unload Timers(Timers.UBound)
    Wend
    
    While Buttons.count > 1
        Unload Buttons(Buttons.UBound)
    Wend
    
    While Images.count > 1
        Unload Images(Images.UBound)
    Wend

    While Labels.count > 1
        Unload Labels(Labels.UBound)
    Wend

End Function

Public Function GoToURL(ByVal theURL As String) As Boolean
    On Error GoTo Handler
    
    ResetWindow

Dim theProtocol As String
Dim thePath As String
Dim strXML As String
    
    theProtocol = Split(theURL, "://")(0)
    thePath = Split(theURL, "://")(1)
    
    If theProtocol = "res" Then
        strXML = StrConv(LoadResData(thePath, "XMLWindow"), vbUnicode, 1033)
    ElseIf theProtocol = "local" Then
        strXML = Fload(thePath)
    Else
        strXML = HttpRequest(theURL)
    End If
    
    If strXML = "" Then
        GoToURL = False
        Exit Function
    End If
    
    If m_creditsXML.loadXML(strXML) Then
        ParseCreditsXML
    End If

    GoToURL = True
    Exit Function
Handler:
    
End Function

Private Sub cmdOK_onClick()
    Unload Me
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "APPICON", False

    Me.BackColor = vbBlack
    m_glassMode = True

    If ApplyGlassIfPossible(Me.hWnd) = False Then
        m_glassMode = False
        Me.BackColor = vbWhite
    End If
    
End Sub

Private Sub Timers_Timer(index As Integer)
Dim theCommand As CommandType

    theCommand = ParseCommandText(Timers(index).Tag)
    ExecuteCommand Me, theCommand.commandText, theCommand.Args
End Sub

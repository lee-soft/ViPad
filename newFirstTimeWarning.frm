VERSION 5.00
Begin VB.Form FirstTimeWarning 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FirstTimeWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Option Explicit

Private m_glassMode As Boolean
Private m_textLayer As ViTextLayer
'Private m_addURL As ViText
Private m_URL As ViText

Private WithEvents m_cmdCancel As ViCommandButton
Attribute m_cmdCancel.VB_VarHelpID = -1
Private WithEvents m_cmdYes As ViCommandButton
Attribute m_cmdYes.VB_VarHelpID = -1
Private WithEvents m_cmdNo As ViCommandButton
Attribute m_cmdNo.VB_VarHelpID = -1

Private m_viComponents As Collection
Private m_backBuffer As GDIPBitmap
Private m_graphicsWindow As GDIPGraphics

Private m_activeObject As Object

Public Event onLoad()
Public Event onCancel()
Public Event onConfirm()

Public AddDesktopShortcuts As Boolean

Public Property Get URL() As String
    URL = m_URL.Caption
End Property

Public Property Let URL(newURL As String)

Dim initialDifference As Long
    initialDifference = (m_URL.CalculateWidth - Me.ScaleWidth) / m_URL.Size

    m_URL.Caption = newURL
    
    If m_URL.CalculateWidth > Me.ScaleWidth - 30 Then
        While (m_URL.CalculateWidth > Me.ScaleWidth - 33)
            m_URL.Caption = Mid(m_URL.Caption, 1, Len(m_URL.Caption) - 1)
        Wend
        m_URL.Caption = m_URL.Caption & "..."
            
    End If
    
    ReDrawComponents
End Property

Private Sub Form_Load()
    InitializeGDIIfNotInitialized
    
    SetIcon Me.hWnd, "APPICON", True
    
    'Set m_glassText = InitializeFormForGlassText(Me)
    Set m_viComponents = New Collection
    
    Set m_backBuffer = New GDIPBitmap
    Set m_graphicsWindow = New GDIPGraphics
    
    Set m_cmdCancel = New ViCommandButton
    Set m_cmdYes = New ViCommandButton
    Set m_cmdNo = New ViCommandButton
    
    Set m_textLayer = New ViTextLayer
    
    m_textLayer.Parent = Me
    
    m_viComponents.Add m_textLayer
    m_viComponents.Add m_cmdNo
    m_viComponents.Add m_cmdYes
    m_viComponents.Add m_cmdCancel
    
    m_cmdCancel.Y = 160
    m_cmdCancel.X = 16
    m_cmdCancel.Width = 109
    m_cmdCancel.Caption = "Cancel"
    
    m_cmdYes.Y = 160
    m_cmdYes.X = 424
    m_cmdYes.Width = 109
    m_cmdYes.Caption = "Yes"
    
    m_cmdNo.Y = 160
    m_cmdNo.X = 544
    m_cmdNo.Width = 109
    m_cmdNo.Caption = "No"
    
    If ApplyGlassIfPossible(Me.hWnd) = False Then
    Else
        m_glassMode = True
    End If
    
    m_textLayer.CreateChild "This looks like the first time you've used ViPad.", _
                            12, _
                            16, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "To help you get started ViPad can automatically add the shortcuts already on your desktop.", _
                            12, _
                            40, _
                            Me.FontName, _
                            Me.fontSize
                            
    m_textLayer.CreateChild "Would you like me to do that now for you?", _
                            12, _
                            112, _
                            Me.FontName, _
                            Me.fontSize

    
    m_graphicsWindow.FromHDC Me.hdc
    m_graphicsWindow.SmoothingMode = SmoothingModeHighQuality
    m_graphicsWindow.InterpolationMode = InterpolationModeHighQualityBicubic
    m_graphicsWindow.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    ReDrawComponents
End Sub

Private Sub ReDrawComponents()
    If m_viComponents Is Nothing Then Exit Sub

    If m_glassMode Then
        m_graphicsWindow.Clear
    Else
        m_graphicsWindow.Clear Me.BackColor
    End If
  
Dim thisObject As Object
    For Each thisObject In m_viComponents
        thisObject.Draw m_graphicsWindow
    Next

    Me.Refresh
End Sub

Private Sub Form_Resize()
    m_graphicsWindow.FromHDC Me.hdc
    m_graphicsWindow.SmoothingMode = SmoothingModeHighQuality
    m_graphicsWindow.InterpolationMode = InterpolationModeHighQualityBicubic
    m_graphicsWindow.TextRenderingHint = TextRenderingHintClearTypeGridFit
    
    ReDrawComponents
End Sub

Private Sub m_cmdCancel_onClicked()
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            thisObject.onMouseDown CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        End If
    Next

    If redrawMe Then
        ReDrawComponents
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            thisObject.onMouseMove CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        End If
    Next

    If redrawMe Then
        ReDrawComponents
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim thisObject As Object
Dim redrawMe As Boolean
Dim Dimensions As RECTL

    For Each thisObject In m_viComponents
        Dimensions = Unserialize_RectL(thisObject.Dimensions_Serialized)
    
        If IsMouseInsideRect(CLng(X), CLng(Y), Dimensions) Then
            thisObject.onMouseUp CLng(Button), X - Dimensions.Left, Y - Dimensions.Top
            
            If thisObject.RedrawRequest Then
                redrawMe = True
            End If
        Else
            If Not m_activeObject Is Nothing Then
                m_activeObject.onMouseOut
                Set m_activeObject = Nothing
            End If
        End If
    Next

    If redrawMe Then
        ReDrawComponents
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_graphicsWindow = Nothing
    Set m_viComponents = Nothing
    Set m_textLayer = Nothing
End Sub

Private Sub m_cmdNo_onClicked()
    AddDesktopShortcuts = False
    Unload Me
End Sub

Private Sub m_cmdYes_onClicked()
    AddDesktopShortcuts = True
    Unload Me
End Sub



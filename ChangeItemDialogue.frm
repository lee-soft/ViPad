VERSION 5.00
Begin VB.Form ChangeItemWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Item Properties"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI Semibold"
      Size            =   9.75
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNewIcons 
      Caption         =   "&Get New Icons"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   540
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowseTarget 
      Caption         =   "..."
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   240
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   11
      Top             =   240
      Width           =   1920
   End
   Begin VB.CommandButton cmdChangeIcon 
      Caption         =   "Change &Icon"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1020
      Width           =   1815
   End
   Begin VB.CommandButton cmdGetFileLocation 
      Caption         =   "Open &File Location"
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtArguements 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox txtTarget 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtCaption 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Arguements:"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Target:"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caption:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "ChangeItemWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'

Option Explicit

Private Const ICONPACK_URL As String = "http://lee-soft.com/vipad/icons"

Private m_graphics As GDIPGraphics
Private m_currentBitmap As GDIPBitmap

Private m_desiredHeight As Long
Private m_desiredWidth As Long
Private m_iconPath As String
Private m_appliedChanges As Boolean
Private m_icon As GDIPImage

Public Property Get IconImage() As GDIPImage
    Set IconImage = m_icon
End Property

Public Property Get Changed() As Boolean
    Changed = m_appliedChanges
End Property

Public Property Let ItemCaption(ByVal newCaption As String)
    txtCaption.Text = newCaption
End Property

Public Property Get ItemCaption() As String
    ItemCaption = txtCaption.Text
End Property

Public Property Get Target() As String
    Target = txtTarget.Text
End Property

Public Property Let Target(ByVal newTarget As String)
    txtTarget.Text = newTarget
End Property

Public Property Get Arguements() As String
    Arguements = txtArguements.Text
End Property

Public Property Let Arguements(ByVal newArguements As String)
    txtArguements.Text = newArguements
End Property

Public Property Get IconPath() As String
    IconPath = Replace(m_iconPath, ApplicationDataPath & "\", "")
End Property

Public Property Let IconPath(ByVal newIconPath As String)
    Set m_icon = New GDIPImage

    m_icon.FromFile newIconPath
    m_iconPath = newIconPath

    m_graphics.Clear vbWhite
    m_graphics.DrawImage m_icon, 0, 0, picIcon.Width, picIcon.Height

    picIcon.Refresh
End Property

Private Sub cmdApply_Click()
    txtArguements.Text = Replace(txtArguements, """", "'")
    
    m_appliedChanges = True
    Me.Hide
End Sub

Private Sub cmdBrowseTarget_Click()
Dim OpenFile As OPENFILENAMEA
Dim lReturn As Long
Dim sFilter As String
Dim theFile As String * 4096

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = g_mainHwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "Programs (*.exe)" & Chr(0) & "*.EXE" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = theFile
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    'OpenFile.lpstrInitialDir = "C:\"
    OpenFile.lpstrTitle = "Select"
    OpenFile.Flags = 0
    lReturn = GetOpenFileNameA(OpenFile)
    
    If Not lReturn = 0 Then
       txtTarget.Text = Trim(OpenFile.lpstrFile)
    End If
End Sub

Private Sub cmdCancel_Click()
    m_appliedChanges = False
    Me.Hide
End Sub

Private Sub cmdChangeIcon_Click()
Dim OpenFile As OPENFILENAMEA
Dim lReturn As Long
Dim sFilter As String
Dim theFile As String * 4096
Dim imageFileNamePath As String
Dim newImage As New GDIPImage

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = g_mainHwnd
    OpenFile.hInstance = App.hInstance

    sFilter = "Portable Network Graphics" & Chr(0) & "*.png" & Chr(0) & _
              "Tagged Image File Format" & Chr(0) & "*.tiff;*.tif" & Chr(0) & _
              "Icon Files" & Chr(0) & "*.ico" & Chr(0)

    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = theFile
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = App.Path & "\icons\"
    OpenFile.lpstrTitle = "Select"
    OpenFile.Flags = 0
    lReturn = GetOpenFileNameA(OpenFile)
    
    If Not lReturn = 0 Then
        newImage.FromFile Trim(OpenFile.lpstrFile)
        imageFileNamePath = GenerateAvailableFileName
        
        If imageFileNamePath <> vbNullString Then
            
            'newImage.Save imageFileNamePath, GetPngCodecCLSID()
            SaveResizedImage newImage, imageFileNamePath
            Me.IconPath = imageFileNamePath
        End If
    End If
End Sub

Private Function SaveResizedImage(ByRef theOriginalImage As GDIPImage, ByVal thePath As String)

Dim theBitmap As New GDIPBitmap
Dim theGraphics As New GDIPGraphics

    theBitmap.CreateFromSizeFormat 256, 256, PixelFormat.Format32bppArgb
    theGraphics.FromImage theBitmap.Image
    theGraphics.Clear
    
    theGraphics.DrawImage theOriginalImage, 0, 0, 256, 256
    theBitmap.Image.Save thePath, GetPngCodecCLSID

End Function

Private Sub cmdGetFileLocation_Click()
 
Dim theFileLocation As String

    theFileLocation = GetFileLocation(txtTarget.Text)
    If theFileLocation = "" Then
       MessageBox Me.hWnd, "The target file doesn't exist!", "Target Error", MB_ICONEXCLAMATION
       Exit Sub
    End If

    On Error Resume Next
    Shell "explorer.exe " & """" & theFileLocation & """", vbNormalFocus
    
End Sub

Private Sub cmdNewIcons_Click()

    On Error GoTo Handler
    
    ShellExecute Me.hWnd, "open", ICONPACK_URL, 0, 0, SW_SHOW
    Exit Sub
Handler:
End Sub

Private Sub Form_Initialize()
    m_desiredHeight = Me.ScaleHeight
    m_desiredWidth = Me.ScaleWidth
End Sub

Private Sub Form_Load()
    On Error GoTo Handler

    SetIcon Me.hWnd, "APPICON", False

Dim negativeRect As RECTL
    negativeRect.Left = -1
    negativeRect.Top = -1
    
    Set m_graphics = New GDIPGraphics
    Set m_currentBitmap = New GDIPBitmap

    m_graphics.FromHDC picIcon.hdc
    m_graphics.SmoothingMode = SmoothingModeHighQuality
    m_graphics.InterpolationMode = InterpolationModeHighQuality
    
    'm_graphicsImage.FromImage m_currentBitmap.Image
    'Win32.ApplyGlass Me.hWnd, negativeRect
    IconPath = m_iconPath
    
    If Config.TopMostWindow Then StayOnTop Me, True
    
    Exit Sub
Handler:
End Sub

Private Sub RefreshHDC()

    'On Error GoTo Handler

    'If m_graphics Is Nothing Then Exit Sub
    
End Sub


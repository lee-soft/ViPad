VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Slice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Anchor As AnchorPointConstants
Public Image As GDIPImage

Public X As Long
Public Y As Long
Public StretchX As Boolean
Public StretchY As Boolean

Public PixelGap As Long
Public PixelGap2 As Long

Public Property Get Height() As Long
   Height = Image.Height
End Property

Public Property Get Width() As Long
   Width = Image.Width
End Property

Public Function LoadSliceFromXMLElement()



End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_itemBank As Collection

Public masterCollection As Collection

Public Function CleanMasterCollectionOfCollection(ByRef theCollection As Collection)
    
Dim thisItem As LaunchPadItem
    For Each thisItem In theCollection
        masterCollection.Remove thisItem.GlobalIdentifer
    Next
End Function

Public Function CleanMasterCollection(ByRef collectionIndex As Long)
    
Dim thisCollection As Collection
Dim thisItem As LaunchPadItem
    
    Set thisCollection = m_itemBank(collectionIndex).Children
    For Each thisItem In thisCollection
        masterCollection.Remove thisItem.GlobalIdentifer
    Next
End Function

Public Function RemoveTab(ByVal theTabIdentifier As String)
    Dim thisCollection As Collection
    Set thisCollection = m_itemBank(theTabIdentifier).Children
    
    CleanMasterCollectionOfCollection thisCollection
    m_itemBank.Remove theTabIdentifier
End Function

Public Function RemoveIndex(ByVal collectionIndex As Long)
    CleanMasterCollection collectionIndex

    m_itemBank.Remove collectionIndex
    'm_itemAliases.Remove collectionIndex
End Function

Public Function GetBank() As Collection
    Set GetBank = m_itemBank
End Function

Public Function GetCollectionCount()
    GetCollectionCount = m_itemBank.count
End Function

Public Function SetCollectionAlias(ByVal collectionIndex As Long, newAlias As String)
    
    m_itemBank(collectionIndex).Alias = newAlias

End Function

Public Function GetTabByIndex(ByVal tabIndex As Long) As ViTab
    On Error GoTo Handler
    
    Set GetTabByIndex = m_itemBank(tabIndex)
    Exit Function
Handler:
    MsgBox "Tab Index: '" & tabIndex & "' doesn't exist!", vbCritical
End Function

Public Function GetCollectionAlias(ByVal collectionIndex As Long) As String
    On Error GoTo Handler
    
    GetCollectionAlias = m_itemBank(collectionIndex).Alias
    Exit Function
Handler:
    MsgBox "Collection Index: '" & collectionIndex & "' doesn't exist!", vbCritical

End Function

Public Function GetCollectionByIndex(ByVal collectionIndex As Long) As Collection
    On Error GoTo Handler
    
    While Me.GetCollectionCount < collectionIndex
        Me.AddNewCollection "New Tab"
    Wend
    
    Set GetCollectionByIndex = m_itemBank(collectionIndex).Children
    Exit Function
Handler:
    MsgBox "Collection Index: '" & collectionIndex & "' doesn't exist!", vbCritical
End Function

Public Function AddNewCollection(ByVal collectionIdentifier As String) As ViTab
    Debug.Print "AddNewCollection!"
    
    Set AddNewCollection = New ViTab
    
    AddNewCollection.Alias = collectionIdentifier
    AddNewCollection.SharedViPadIdentifer = ProgramSupport.GetNextViPadKey
    
    m_itemBank.Add AddNewCollection, AddNewCollection.SharedViPadIdentifer

    PrintHeader Me
End Function

Private Sub Class_Initialize()
    Set masterCollection = New Collection
    Set m_itemBank = New Collection
    'Set m_itemAliases = New Collection
End Sub

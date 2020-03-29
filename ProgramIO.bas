Attribute VB_Name = "ProgramIO"
Option Explicit

Private m_sourceDoc As DOMDocument

Private Function LoadXMLPad(theBank As ViBank, ByRef XMLPad As IXMLDOMElement, targetCollection As Collection)

Dim thisPadItemXML As IXMLDOMElement
Dim thisPadItem As LaunchPadItem

Dim thisPadCollection As Collection

Dim thisIconImage As GDIPImage
Dim thisAlphaIcon As AlphaIcon

    For Each thisPadItemXML In XMLPad.selectNodes("item")
    
        Debug.Print thisPadItemXML.tagName
    
        Set thisPadItem = New LaunchPadItem
        Set thisPadItem.Icon = New GDIPImage

        thisPadItem.TargetPath = thisPadItemXML.getAttribute("target")
        thisPadItem.IconPath = thisPadItemXML.getAttribute("icon")

        thisPadItem.Caption = thisPadItemXML.getAttribute("caption")
        thisPadItem.Icon.FromFile ApplicationDataPath & "\" & thisPadItem.IconPath
        
        
        thisPadItem.TargetArguements = thisPadItemXML.getAttribute("arguements")
        
        'AddLaunchItem thisPadItem
        targetCollection.Add thisPadItem, thisPadItem.GlobalIdentifer
        theBank.masterCollection.Add thisPadItem, thisPadItem.GlobalIdentifer
    Next

End Function

Public Function LoadBankXML(szPath As String) As ViBank

Dim padItems As DOMDocument
Dim thisPadItemXML As IXMLDOMElement

Dim thisPadItem As LaunchPadItem

Dim thisPadCollection As Collection

Dim thisIconImage As GDIPImage
Dim thisAlphaIcon As AlphaIcon

Dim padAlias As String
Dim newBank As ViBank
Dim windowCount As Long
Dim windowDimensions As GdiPlus.RECTL

    Set newBank = New ViBank
    Set LoadBankXML = newBank

    Set padItems = New DOMDocument
    
    If padItems.Load(szPath) = False Then
        Exit Function
    End If
    
    For Each thisPadItemXML In padItems.selectNodes("pads/pad")
        padAlias = thisPadItemXML.getAttribute("alias")

        If Not IsNull(thisPadItemXML.getAttribute("height")) Then windowDimensions.Height = thisPadItemXML.getAttribute("height")
        If Not IsNull(thisPadItemXML.getAttribute("width")) Then windowDimensions.Width = thisPadItemXML.getAttribute("width")
        If Not IsNull(thisPadItemXML.getAttribute("top")) Then windowDimensions.Top = thisPadItemXML.getAttribute("top")
        If Not IsNull(thisPadItemXML.getAttribute("left")) Then windowDimensions.Left = thisPadItemXML.getAttribute("left")
        
        With newBank.AddNewCollection(padAlias)
            .Dimensions = Serialize_RectL(windowDimensions)
        End With
        
        LoadXMLPad newBank, thisPadItemXML, newBank.GetCollectionByIndex(newBank.GetCollectionCount)
        windowCount = windowCount + 1
    Next

End Function

Public Function GetPadsXMLDocument(ByRef itemBank As ViBank) As DOMDocument

Dim rootElement As IXMLDOMElement

Dim thisCollection As Collection
Dim thisTab As ViTab

Dim thisWindow As ViPickWindow
Dim collectionIndex As Long

    Set m_sourceDoc = New DOMDocument
    
    Set rootElement = m_sourceDoc.createElement("pads")
    m_sourceDoc.appendChild rootElement
    
    For collectionIndex = 1 To itemBank.GetCollectionCount
        Set thisTab = itemBank.GetTabByIndex(collectionIndex)
        rootElement.appendChild GetPadXMLElement(thisTab.Children, thisTab.Alias, thisTab.Dimensions)
    Next

    Set GetPadsXMLDocument = m_sourceDoc
End Function

Private Function GetPadXMLElement(ByRef Items As Collection, ByVal collectionId As String, Optional szWindowDimensions As String) As IXMLDOMElement

Dim szXML As String
Dim thisLaunchPadItem As LaunchPadItem

Dim newPad As IXMLDOMElement
Dim newItem As IXMLDOMElement
Dim windowDimensions As RECTL
    
    Set newPad = m_sourceDoc.createElement("pad")
    newPad.setAttribute "alias", collectionId
    
    If Not szWindowDimensions = vbNullString Then
        windowDimensions = Unserialize_RectL(szWindowDimensions)

        newPad.setAttribute "left", windowDimensions.Left
        newPad.setAttribute "top", windowDimensions.Top
        newPad.setAttribute "width", windowDimensions.Width
        newPad.setAttribute "height", windowDimensions.Height
    End If

    For Each thisLaunchPadItem In Items
        Set newItem = m_sourceDoc.createElement("item")
        newPad.appendChild newItem
        
        newItem.setAttribute "icon", thisLaunchPadItem.IconPath
        newItem.setAttribute "caption", thisLaunchPadItem.Caption
        newItem.setAttribute "target", thisLaunchPadItem.TargetPath
        newItem.setAttribute "arguements", thisLaunchPadItem.TargetArguements
    Next

    Set GetPadXMLElement = newPad
End Function


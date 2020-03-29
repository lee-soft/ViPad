Attribute VB_Name = "XMLParser"
Option Explicit

Private Const XML_IDENTIFIER As String = "XML_"
Private Const ATTRIB_IDENTIFIER As String = "ATTRIB_"

Public XMLElementCount As Long
Private m_attributeCount As Long

Public Function GetNextElementID(Optional elementCollection As Collection)
    If XMLElementCount < 2147483647 Then
        XMLElementCount = XMLElementCount + 1
    Else
        XMLElementCount = 0
    End If

    GetNextElementID = XML_IDENTIFIER & XMLElementCount
End Function

Public Function GetNextAttributeID(Optional elementCollection As Collection)
    If m_attributeCount < 2147483647 Then
        m_attributeCount = m_attributeCount + 1
    Else
        m_attributeCount = 0
    End If

    GetNextAttributeID = ATTRIB_IDENTIFIER & m_attributeCount
End Function

Public Function MakeAttribute(theName As String, theValue As String) As XMLAttribute

Dim attrib As New XMLAttribute

    With attrib
        .Name = theName
        .Value = theValue
    End With
    
    Set MakeAttribute = attrib
End Function

Public Function ParseAttributes(ByRef attributeCollection As Collection, sourceXMLHeader As String)

Dim nextWhiteSpace As Long
Dim thisAttribute As String
Dim thisAttributeObj As XMLAttribute
Dim equalPosition As Long

Dim FirstQuote As Long
Dim SecondQuote As Long

    sourceXMLHeader = Replace(sourceXMLHeader, vbCrLf, "")

    While Len(sourceXMLHeader) > 0
        sourceXMLHeader = Trim(sourceXMLHeader)
        equalPosition = InStr(sourceXMLHeader, "=")
        
        FirstQuote = InStr(sourceXMLHeader, """") + 1
        SecondQuote = InStr(FirstQuote, sourceXMLHeader, """")
        
        If FirstQuote = 1 Or SecondQuote = 0 Then Exit Function
        
        Set thisAttributeObj = New XMLAttribute
        
        thisAttributeObj.Name = Trim(Mid(sourceXMLHeader, 1, equalPosition - 1))
        thisAttributeObj.Value = Trim(Mid(sourceXMLHeader, FirstQuote, SecondQuote - FirstQuote))
        
        attributeCollection.Add thisAttributeObj, thisAttributeObj.Key
        sourceXMLHeader = Mid(sourceXMLHeader, SecondQuote + 1)
    Wend

End Function

Private Function GetLastIncompleteType(ByRef xmlHeaders As Collection, ElementType As String) As XMLFragment

Dim thisElement As XMLFragment
Dim elementIndex As Long

    If xmlHeaders.Count = 0 Then
        Exit Function
    End If

    For elementIndex = 0 To xmlHeaders.Count - 1
        Set thisElement = xmlHeaders(xmlHeaders.Count - elementIndex)

        'Debug.Print thisElement.ElementType

        If Not thisElement.Complete And thisElement.ElementType = ElementType Then
            Set GetLastIncompleteType = thisElement
            Exit For
        End If
    Next
End Function

Private Function GetLastIncomplete(ByRef xmlHeaders As Collection) As XMLFragment

Dim thisElement As XMLFragment
Dim elementIndex As Long

    If xmlHeaders.Count = 0 Then
        Exit Function
    End If

    For elementIndex = 0 To xmlHeaders.Count - 1
        Set thisElement = xmlHeaders(xmlHeaders.Count - elementIndex)

        If Not thisElement.Complete Then
            Set GetLastIncomplete = thisElement
            Exit For
        End If
    Next

End Function

Public Function ParseFragmentedXML(ByRef xmlHeaders As Collection, ByRef m_buffer As String)

Dim charIndex As Long
Dim inXmlElementHeader As Boolean
Dim inXmlElementFooter As Boolean
Dim inTypeDefinition As Boolean

Dim inQuote As Boolean
Dim thisXML As XMLFragment
Dim thisChar As String
Dim nextChar As String

Dim oldXML As XMLFragment
Dim footerType As String
Dim headerIndex As Long

Dim lastIncompleteElement As XMLFragment
Dim cutPoint As Long
Dim tempStr As String

    Set thisXML = New XMLFragment
    thisXML.Key = GetNextElementID()

    charIndex = 1
    thisXML.Orphan = True
    
    While charIndex <= Len(m_buffer)
        thisChar = Mid(m_buffer, charIndex, 1)
        If thisChar = """" Then
            If inQuote Then
                inQuote = False
            Else
                inQuote = True
            End If
        End If
        
        Set lastIncompleteElement = GetLastIncomplete(xmlHeaders)
        thisXML.XML = thisXML.XML & thisChar

        If Not lastIncompleteElement Is Nothing Then
            If Not inXmlElementHeader And Not inXmlElementFooter And _
                Not thisChar = "<" And Not thisChar = ">" Then
                
                lastIncompleteElement.Text = lastIncompleteElement.Text & thisChar
                lastIncompleteElement.XML = lastIncompleteElement.XML & thisChar
                
                If charIndex > cutPoint Then
                    cutPoint = charIndex
                End If
            End If
        End If
            
        If Not inQuote Then
            If inTypeDefinition Then
                If thisChar = " " Or thisChar = ">" Then
                    inTypeDefinition = False
                End If
            End If
            
            If thisChar = "/" Then
                If inXmlElementHeader Then
                    If charIndex < Len(m_buffer) Then
                        nextChar = Mid(m_buffer, charIndex + 1, 1)
                        inTypeDefinition = False
                        inXmlElementFooter = True
                        
                        If Not nextChar = ">" Then
                            inXmlElementHeader = False
                        Else
                            footerType = thisXML.ElementType
                        End If
                    End If
                End If
            End If
        
            If inTypeDefinition Then
                thisXML.ElementType = thisXML.ElementType & thisChar
            End If
            
            If inXmlElementFooter Then
                If thisChar = " " Or thisChar = ">" Then
                    inXmlElementFooter = False

                    'Add support for ShortTaggedXML <a/>
                    If footerType = thisXML.ElementType Then
                        Set oldXML = thisXML
                    Else
                        Set oldXML = GetLastIncompleteType(xmlHeaders, footerType)
                        oldXML.XML = oldXML.XML & "</" & footerType & ">"
                    End If
                    
                    If Not oldXML Is Nothing Then
                    
                        oldXML.Complete = True
                        If Not oldXML.Parent Is Nothing Then
                            tempStr = "<#" & oldXML.Key & "#>"
                            oldXML.Parent.XML = Replace(oldXML.Parent.XML, tempStr, oldXML.XML, 1, 1)
                            
                            oldXML.Parent.children.Add oldXML, oldXML.Key
                        End If
                        
                        If charIndex > cutPoint Then
                            cutPoint = charIndex
                        End If
                        
                        footerType = ""
                    End If
                ElseIf Not thisChar = "/" Then
                    footerType = footerType & thisChar
                End If
            End If
            
            'Start XML Element Definition
            If Not inXmlElementHeader Then
                If thisChar = "<" Then
                    If charIndex < Len(m_buffer) Then
                        nextChar = Mid(m_buffer, charIndex + 1, 1)
                        
                        'Beginning a new element header definately terminates within the buffer
                        If Not nextChar = "/" And InStr(charIndex, m_buffer, ">") > 0 Then
                            Set oldXML = GetLastIncomplete(xmlHeaders)

                            If Not oldXML Is Nothing Then
                                thisXML.Orphan = False
                                
                                Set thisXML.Parent = oldXML
                                tempStr = "<#" & thisXML.Key & "#>"
                                
                                If Not Right(oldXML.XML, Len(tempStr)) = tempStr Then
                                    oldXML.XML = oldXML.XML & tempStr
                                End If
                            End If
                        
                            thisXML.XML = "<"
                        End If
                    
                        inXmlElementHeader = True
                        inTypeDefinition = True
                    End If
                End If
            Else
                If thisChar = ">" Then
                    inXmlElementHeader = False
                    xmlHeaders.Add thisXML, thisXML.Key

                    Set thisXML = New XMLFragment
                    thisXML.Key = GetNextElementID(xmlHeaders)
                    
                    thisXML.Orphan = True
                    
                    If charIndex > cutPoint Then
                        cutPoint = charIndex
                    End If
                End If
            End If
        End If
        
        charIndex = charIndex + 1
    Wend
    
    If cutPoint > 0 Then
        m_buffer = Mid(m_buffer, cutPoint + 1)
    End If
    
End Function

Public Function ParseXML(ByRef targetXML As XMLElement2, sourceXML As String)

Dim endHeader As Long
Dim startHeader As Long
Dim allInclusive As Long
Dim firstWhiteSpace As Long
Dim elementTypeDef As String

    If targetXML Is Nothing Then
        Exit Function
    End If
    If sourceXML = "" Then Exit Function

Dim thisXML As XMLElement2

    endHeader = InStr(sourceXML, ">")
    startHeader = InStr(sourceXML, "<")
    
    allInclusive = InStr(sourceXML, "/>")
    firstWhiteSpace = InStr(sourceXML, " ")
    If firstWhiteSpace = 0 Then firstWhiteSpace = InStr(sourceXML, vbCrLf)
    'No More XML
    If startHeader = 0 Then Exit Function
    
    If InStr(sourceXML, "</") = 1 Then
        sourceXML = Mid(sourceXML, endHeader + 1)
        
        Set thisXML = New XMLElement2
        Set thisXML.Parent = targetXML.Parent.Parent

        ParseXML thisXML, sourceXML
        Exit Function
    End If
    
    If Not startHeader = 1 Then
        Set thisXML = New XMLElement2
        thisXML.ElementType = ""
        Set thisXML.Parent = targetXML.Parent

        targetXML.XML = Mid(sourceXML, 1, startHeader - 1)
        If Not targetXML.Parent Is Nothing Then
            targetXML.Parent.AddChild targetXML
        End If
        
        sourceXML = Mid(sourceXML, startHeader)
        ParseXML thisXML, sourceXML
        Exit Function
    End If
    
    If firstWhiteSpace < endHeader And firstWhiteSpace > 0 Then
        elementTypeDef = Mid(sourceXML, 2, firstWhiteSpace - 2)
        
        If allInclusive > 0 And allInclusive < endHeader Then
            targetXML.Header = Mid(sourceXML, firstWhiteSpace, allInclusive - firstWhiteSpace)
        Else
            targetXML.Header = Mid(sourceXML, firstWhiteSpace, endHeader - firstWhiteSpace)
        End If
    Else
        elementTypeDef = Replace(Mid(sourceXML, 2, endHeader - 2), "/", "")
    End If
    
    targetXML.ElementType = elementTypeDef

    'All inclusive XML element, no children!
    If allInclusive > 0 And allInclusive < endHeader Then
        sourceXML = Mid(sourceXML, allInclusive + 2)
        
        Set thisXML = New XMLElement2
        Set thisXML.Parent = targetXML.Parent

        'Incase its a standalone inclusive xml element
        If Not targetXML.Parent Is Nothing Then
            targetXML.Parent.AddChild targetXML
        End If
        
        ParseXML thisXML, sourceXML
    Else
        Set thisXML = New XMLElement2
        Set thisXML.Parent = targetXML
        
        If Not targetXML.Parent Is Nothing Then
            targetXML.Parent.AddChild targetXML
        End If
    
        sourceXML = Mid(sourceXML, endHeader + 1)
        ParseXML thisXML, sourceXML
    End If
        

End Function

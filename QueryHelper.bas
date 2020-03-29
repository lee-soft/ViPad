Attribute VB_Name = "QueryHelper"
Private Function IsValidArray(arr As Variant) As Boolean
    On Error Resume Next
    
    If Not UBound(arr) Then
        IsValidArray = True
    End If
End Function

Public Function QueryByString(ByRef masterCollection As Collection, Optional ByVal theName As String) As Collection
    On Error GoTo Handler
    
Dim returnCollection As New Collection
Dim thisNodeIndex As Long
Dim thisNode As LaunchPadItem

Dim keyWords() As String
Dim lenKeywords() As Long
Dim keyWordIndex As Long

Dim thisWordSpacePositions() As Long
Dim thisItemName As String
Dim thisSpaceIndex As Long
Dim thisKeyWordIndex As Long

Dim bShow() As Boolean
Dim bShowFinal As Boolean

    Set QueryByString = returnCollection

    If masterCollection Is Nothing Then
        Exit Function
    End If
    
    If Len(theName) = 0 Then
        Exit Function
    End If
    
    keyWords = Split(UCase(theName), " ")
    ReDim lenKeywords(UBound(keyWords))
    
    'Get the length of each word
    For keyWordIndex = 0 To UBound(keyWords)
        lenKeywords(keyWordIndex) = Len(keyWords(keyWordIndex))
    Next
    
    For Each thisNode In masterCollection
        'If IsObject(masterCollection(thisNodeIndex)) Then
        'Set thisNode = masterCollection(thisNodeIndex)

        If Not thisNode Is Nothing Then
            added = False
            thisWordSpacePositions = thisNode.SpacePositions
            
            If IsValidArray(thisWordSpacePositions) Then
                thisItemName = thisNode.SearchIdentifier
                
                ReDim bShow(UBound(keyWords))
                bShowFinal = True
                
                'For Each First Letter
                For thisSpaceIndex = 0 To UBound(thisWordSpacePositions)
        
                    For thisKeyWordIndex = 0 To UBound(keyWords)
    
                        If Mid(thisItemName, thisNode.SpacePositions(thisSpaceIndex), lenKeywords(thisKeyWordIndex)) = keyWords(thisKeyWordIndex) Then
                            bShow(thisKeyWordIndex) = True
                        End If
                    Next
                    
                Next
                
                For thisKeyWordIndex = LBound(bShow) To UBound(bShow)
                    If bShow(thisKeyWordIndex) = False Then
                        bShowFinal = False
                        Exit For
                    End If
                Next
                
                Debug.Print "Found:: " & thisNode.Caption
                If bShowFinal Then returnCollection.Add thisNode, thisNode.GlobalIdentifer
            End If
        End If
        'Else
            'MsgBox TypeName(masterCollection(thisNodeIndex))
        'End If
    Next
    
Handler:
End Function

Public Function GetSpacePositions(theWord As String) As Long()

Dim mvarSpacePositions() As Long
Dim lngSpaceCount As Long
Dim charIndex As Long

    ReDim mvarSpacePositions(lngSpaceCount)
    mvarSpacePositions(0) = 1
    lngSpaceCount = 1

    For charIndex = 1 To Len(theWord)
        If Asc(Mid(theWord, charIndex, 1)) = vbKeySpace Then
            ReDim Preserve mvarSpacePositions(lngSpaceCount)
            mvarSpacePositions(lngSpaceCount) = charIndex + 1
            
            lngSpaceCount = lngSpaceCount + 1
        End If
    Next

    GetSpacePositions = mvarSpacePositions

End Function

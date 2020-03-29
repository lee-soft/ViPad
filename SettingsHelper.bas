Attribute VB_Name = "SettingsHelper"
Option Explicit

Public Function BoolToXML(sourceBool As Boolean) As Long
    If sourceBool = True Then
        BoolToXML = 1
    Else
        BoolToXML = 0
    End If
End Function

Public Function XMLRectToRealRect(ByVal thisRect As IXMLDOMElement) As GdiPlus.RECTL

    With XMLRectToRealRect
        .Left = thisRect.getAttribute("left")
        .Top = thisRect.getAttribute("top")
        .Width = thisRect.getAttribute("width")
        .Height = thisRect.getAttribute("height")
    End With

End Function

Public Function RealRectToXmlString(ByRef thisRect As GdiPlus.RECTL, ByVal thisRectID As String) As String

Dim returnStr As String
    returnStr = "<Rect id=" & """" & thisRectID & """"

    With thisRect
        returnStr = returnStr & " left=" & """" & .Left & """" & _
                                " top=" & """" & .Top & """" & _
                                " width=" & """" & .Width & """" & _
                                " height=" & """" & .Height & """"
    End With
    
    returnStr = returnStr & " />"
    RealRectToXmlString = returnStr

End Function


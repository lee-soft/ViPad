Attribute VB_Name = "EditorSupport"
Public Function OpenFileInEditor(ByVal theFile As Scripting.File)

Dim thisDoc As JSONDocument
Dim thisViewer As frmViewer
Dim theURL As String
    
    Set thisDoc = New JSONDocument
    Set thisViewer = New frmViewer
    
    thisDoc.LoadJson theFile.OpenAsTextStream(ForReading).ReadAll
    theURL = thisDoc.GetProperty("url")
    
    thisViewer.URL = theURL
    thisViewer.Show
    
End Function

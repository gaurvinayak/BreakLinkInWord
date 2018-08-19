filelocation="C:\Users\gaurvin\FolderContainingLink"
saveLocation="C:\Users\gaurvin\FolderWhereYouWantToSave"

Set objFSO=CreateObject("Scripting.FileSystemObject")
objStartFolder=filelocation
Set objFolder=objFSO.GetFolder(objStartFolder)
Set colFiles=objFolder.Files

For each objFile in colFiles
    Set oWord=CreateObject("Word.Application")
    oWord.Visible=True
    oWord.Documents.Open filelocation & objFile.Name
    Set activeDoc=oWord.ActiveDocument
    Call theTrick
    oWord.Quit
Next


Sub theTrick()
    Call saveAsDoc
    Call breakLinks
    activeDoc.save
    oWord.DisplayAlerts=True
End Sub

Sub saveAsDoc()
    Dim newName
    newName= saveLocation & activeDoc.Name
    activeDoc.SaveAs2 newName,wdFormatDocument
End Sub

Sub breakLinks()
For each objField in activeDoc.InLineShapes
    If Not objField.LinkFormat Is nothing Then
        objField.LinkFormat.Update
        objField.LinkFormat.breakLink
        activeDoc.UndoClear
    End if
Next
End Sub
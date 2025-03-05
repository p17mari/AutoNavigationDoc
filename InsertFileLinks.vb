Sub InsertFileLinks()
    ' Declare variables for document path, file names, file paths, and file system objects
    Dim docPath As String
    Dim fileName As String
    Dim filePath As String
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim rng As Range

    ' Ensure the document is saved
    ' If the document is not saved (Path is empty), show a message and exit the macro
    If ActiveDocument.Path = "" Then
        MsgBox "Please save the document first before running this macro.", vbExclamation, "Save Required"
        Exit Sub
    End If

    ' Get the folder where the current Word document is saved
    docPath = ActiveDocument.Path  ' Store the current document's folder path in docPath

    ' Check if the folder exists
    ' Create a FileSystemObject to interact with the file system
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder path exists, if not, show an error message and exit
    If Not objFSO.FolderExists(docPath) Then
        MsgBox "The document folder was not found!", vbCritical, "Error"
        Exit Sub
    End If

    ' Get the folder object for further use
    Set objFolder = objFSO.GetFolder(docPath)  ' objFolder now refers to the folder containing the document

    ' Set range for inserting links in the Word document
    Set rng = ActiveDocument.Range  ' Define the range where we want to insert the hyperlinks
    rng.Collapse wdCollapseEnd  ' Collapse the range to the end, so links are inserted at the end of the document

    ' Loop through each file in the folder
    ' Iterate over each file in the folder where the document is saved
    For Each objFile In objFolder.Files
        fileName = objFile.Name  ' Get the name of the current file
        filePath = docPath & "\" & fileName  ' Construct the full file path (document folder + file name)
        
        ' Insert the file name as a hyperlink in the document
        rng.InsertAfter fileName  ' Insert the file name at the current range position
        rng.Collapse wdCollapseEnd  ' Move the range to the end again after inserting the file name
        ActiveDocument.Hyperlinks.Add Anchor:=rng, Address:=filePath, TextToDisplay:=fileName  ' Create a hyperlink at the range position, linking to the full file path
        
        ' Move to the next line after inserting the hyperlink
        rng.InsertParagraphAfter  ' Insert a new paragraph (empty line) after the hyperlink
        rng.Collapse wdCollapseEnd  ' Collapse the range to the end to prepare for the next file link
    Next objFile

    ' Clean up object variables to free memory
    Set objFSO = Nothing  ' Release the FileSystemObject
    Set objFolder = Nothing  ' Release the folder object
    Set rng = Nothing  ' Release the range object
    
    ' Inform the user that the task has been completed
    MsgBox "List of files has been inserted successfully!", vbInformation, "Done"
End Sub

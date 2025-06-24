Sub InsertFileLinks()
    Dim docPath As String
    Dim objFSO As Object
    Dim objFolder As Object
    Dim rng As Range
    Dim tbl As Table
    Dim rowIndex As Integer

    ' Ensure the document is saved
    If ActiveDocument.Path = "" Then
        MsgBox "Please save the document first before running this macro.", vbExclamation, "Save Required"
        Exit Sub
    End If

    ' Get the folder where the current Word document is saved
    docPath = ActiveDocument.Path

    ' Check if the folder exists
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(docPath) Then
        MsgBox "The document folder was not found!", vbCritical, "Error"
        Exit Sub
    End If

    Set objFolder = objFSO.GetFolder(docPath)

    ' Create a new table with headers
    Set rng = ActiveDocument.Range
    Set tbl = ActiveDocument.Tables.Add(rng, 1, 2)
    tbl.Borders.Enable = True
    tbl.Cell(1, 1).Range.Text = "File Path"
    tbl.Cell(1, 2).Range.Text = "Link"
    rowIndex = 1

    ' Start recursive traversal
    ListFilesRecursive objFolder, tbl, rowIndex

    ' Clean up
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set rng = Nothing

    MsgBox "List of files has been inserted successfully!", vbInformation, "Done"
End Sub

Sub ListFilesRecursive(ByVal folder As Object, ByRef tbl As Table, ByRef rowIndex As Integer)
    Dim subFolder As Object
    Dim objFile As Object

    ' List all files in the current folder
    For Each objFile In folder.Files
        rowIndex = rowIndex + 1
        tbl.Rows.Add

        tbl.Cell(rowIndex, 1).Range.Text = objFile.Path

        tbl.Cell(rowIndex, 2).Range.Hyperlinks.Add _
            Anchor:=tbl.Cell(rowIndex, 2).Range, _
            Address:=objFile.Path, _
            TextToDisplay:="Open"
    Next objFile

    ' Recursively call the function for each subfolder
    For Each subFolder In folder.Subfolders
        ListFilesRecursive subFolder, tbl, rowIndex
    Next subFolder
End Sub

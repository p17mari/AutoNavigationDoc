Sub InsertFileLinks()
    Dim docPath As String
    Dim fileName As String
    Dim filePath As String
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim subFolder As Object
    Dim rng As Range
    Dim folderName As String
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

    ' Get the folder object
    Set objFolder = objFSO.GetFolder(docPath)

    ' Insert a table with two columns (File Name | Link)
    Set rng = ActiveDocument.Range
    Set tbl = ActiveDocument.Tables.Add(rng, 1, 2) ' Start with one row, will add more dynamically
    tbl.Borders.Enable = False ' Remove table borders for a cleaner look
    rowIndex = 1 ' Start at the first row

    ' Loop through each subfolder in the folder
    For Each subFolder In objFolder.Subfolders
        ' Insert folder name as a header row
        tbl.Rows.Add
        rowIndex = rowIndex + 1
        tbl.Cell(rowIndex, 1).Range.Text = subFolder.Name ' Folder name in first column
        tbl.Cell(rowIndex, 1).Range.Bold = True ' Make the folder name bold
        tbl.Cell(rowIndex, 2).Range.Text = "" ' Empty second column for spacing

        ' Loop through each file in the subfolder
        For Each objFile In subFolder.Files
            tbl.Rows.Add
            rowIndex = rowIndex + 1
            fileName = objFSO.GetBaseName(objFile.Name) ' Get file name without extension
            filePath = subFolder.Path & "\" & objFile.Name

            ' Insert file name in the first column
            tbl.Cell(rowIndex, 1).Range.Text = fileName

            ' Insert file name as a hyperlink in the second column
            tbl.Cell(rowIndex, 2).Range.Hyperlinks.Add _
                Anchor:=tbl.Cell(rowIndex, 2).Range, _
                Address:=filePath, _
                TextToDisplay:="Open"
        Next objFile
    Next subFolder

    ' Loop through each file in the main folder (not in subfolders)
    For Each objFile In objFolder.Files
        fileName = objFSO.GetBaseName(objFile.Name) ' Get file name without extension
        filePath = objFile.Path

        ' Insert new row
        tbl.Rows.Add
        rowIndex = rowIndex + 1

        ' Insert file name in the first column
        tbl.Cell(rowIndex, 1).Range.Text = fileName

        ' Insert file name as a hyperlink in the second column
        tbl.Cell(rowIndex, 2).Range.Hyperlinks.Add _
            Anchor:=tbl.Cell(rowIndex, 2).Range, _
            Address:=filePath, _
            TextToDisplay:="Open"
    Next objFile

    ' Clean up
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set subFolder = Nothing
    Set rng = Nothing

    MsgBox "List of files has been inserted successfully!", vbInformation, "Done"
End Sub

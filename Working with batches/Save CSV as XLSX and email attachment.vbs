' Create a FileSystemObject to handle file operations
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Declare paths and email list
Dim csvPath, xlsxPath
emailList = "John@gmail.com;Bill.Gates@gmail.com"

' Get the current batch and the first document in it
Set Batch = ChronoApp.GetCurrentBatch()
Set Doc = Batch.GetDocument(0)

' Retrieve the export info path from the document metadata
expInfo = Doc.get_field_value("expinfo.TXT Report.File System")

' Assign the CSV path from the export info
csvPath = expInfo
'msgbox "CSV " & csvPath  ' Optional debug message

' Extract base name and folder path from the CSV file path
baseName = fso.GetBaseName(csvPath)
folderPath = fso.GetParentFolderName(csvPath)

' Construct the XLSX path by replacing the extension
xlsxPath = folderPath & "\" & baseName & ".xlsx"
'msgbox "XLS " & xlsxPath  ' Optional debug message

' Get the folder containing the CSV file
Dim oFldrPath, oFldr
oFldrPath = fso.GetParentFolderName(expInfo)
Set oFldr = fso.GetFolder(oFldrPath)

' Loop through all files in the folder
For Each ofile In oFldr.Files
    ' Check if the file has a .csv extension (case-insensitive)
    If LCase(fso.GetExtensionName(ofile.Name)) = "csv" Then
        'msgbox xlsxPath  ' Optional debug message

        ' Create an instance of Excel
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False  ' Keep Excel hidden

        ' Open the CSV file in Excel
        Set xlBook = xlApp.Workbooks.Open(ofile.Path)

        ' If the XLSX file already exists, delete it
        If fso.FileExists(xlsxPath) Then
            fso.DeleteFile xlsxPath, True
        End If

        ' Save the workbook as XLSX (format code 51)
        xlBook.SaveAs xlsxPath, 51

        ' Close the workbook without saving changes
        xlBook.Close False

        ' Quit Excel and release objects
        xlApp.Quit
        Set xlBook = Nothing
        Set xlApp = Nothing

        ' Send the XLSX file as an email attachment
        Call ChronoApp.SendSmtpEmail(FromEmail, FromName, emailList, CC, BCC, Subject, Body, xlsxPath)

        ' Delete the original CSV file if it exists
        Call DeleteIfExists(csvPath)
    End If
Next

' Function to delete a file if it exists, with error handling
Function DeleteIfExists(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        On Error Resume Next
        fso.DeleteFile filePath, True
        If Err.Number <> 0 Then
            Msgbox "Error deleting file: " & Err.Description
            Err.Clear
        Else
            ChronoApp.AddToOutputWindow "File deleted: " & filePath
        End If
        On Error GoTo 0
    Else
        ChronoApp.AddToOutputWindow "File does not exist: " & filePath
    End If

    Set fso = Nothing
End Function

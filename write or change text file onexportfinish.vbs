Const ForReading = 1
Const ForWriting = 2

Set Batch=ChronoApp.GetCurrentBatch
'Get path of existing text file
txtpathandFile=Batch.GetSystemField(0,"expinfo.TXT Report.TXT Report to File System")
tmpTextfile=Replace(txtpathandFile,".CSV",".TMP")

'msgbox txtpathandFile

' Create fso
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(txtpathandFile, ForReading)

' Write header data into text file
Set objFile = objFSO.OpenTextFile(tmpTextfile, ForWriting)
Hdr1=Chr(34) & "FORMAT BATCH IMPORT" & Chr(34) & ", STANDARD 1.0,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
objFile.WriteLine Hdr1

 'Read existing text file
strText = objTextFile.ReadAll
objTextFile.Close

' Split text lines to rows
arrLines = Split(strText, vbCrLf)
 

' Manipulate text here and write lines to temp file. If you need i.e. to delete first/last line change start end points of loop
For i = 1 to (Ubound(arrLines) - 1)
    objFile.WriteLine arrLines(i)
Next
 
objFile.Close

' If you need to move, copy or delete the files use the following methods
' objFSO.CopyFile "C:\temp\*.txt", "C:\Windows\Desktop"
' objFSO.MoveFile "C:\temp\*.txt", "C:\Windows\Desktop"
' objFSO.DeleteFile "C:\temp\*.txt", False
   

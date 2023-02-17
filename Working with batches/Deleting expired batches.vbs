'Recurse working folder
WorkDir="C:\ProgramData\Chronoscan\WorkDir"

Set oFSO   = CreateObject("Scripting.FileSystemObject") 
sFileName  = "batch_config.xml"

Set oFolder = oFSO.GetFolder(WorkDir)
Recurse(oFolder)

Sub Recurse(oFldr)
    If IsAccessible(oFolder) Then

        For Each oSubFolder In oFldr.SubFolders
             Recurse oSubFolder
        Next 

        For Each oFile In oFldr.Files
            If LCase(oFile.Name) = sFileName Then
            dateLastmod=CDATE(oFile.DateLastModified)

            noOfdays=DateDiff("d",dateLastmod,Now)
            Set xmlDoc =CreateObject("Microsoft.XMLDOM")

            xmlDoc.Async = "False"
            xmlDoc.Load(oFile.Path)
                
            Set Root=xmlDoc.documentElement
            Set NodeList = xmlDoc.getElementsByTagName("batch_config")

            ' Get all required data from the batch config file
            isProc = xmlDoc.getElementsByTagName("batch_isprocessed").Item(0).text
            numErrors = xmlDoc.getElementsByTagName("batch_num_errors").Item(0).text
            isExported= xmlDoc.getElementsByTagName("batch_exported").Item(0).text
            batchName = xmlDoc.getElementsByTagName("batch_name").Item(0).text
			
			' Uncomment lines below if you need to examine the returned data 	
            'msgbox  " Batch Folder " & oFldr.Path & vbCrLf & " isproc " & isProc & vbCrLf &_
            '" numErrors " & numErrors & vbCrLf & " isExported " & isExported & vbCrLf & " batchName " & batchName & vbCrLf &_
            '" No of days " &  noOfdays
            
            ' Delete folder if spec as below 
            If isProc=1 AND numErrors=0 AND isExported=1 AND noOfdays > 30 Then
				' Uncomment next line after testing, the msgbox shows what WILL be deleted
                MsgBox "Deleting " & oFldr.Path
				'Uncomment line below to enable delete after fully testing your conditions
				'oFSO.DeleteFolder(oFldr.Path)
            End If      


        End If    
        Next 
    End If
End Sub

Function IsAccessible(oFolder)
  On Error Resume Next
  IsAccessible = (oFolder.SubFolders.Count >= 0)
End Function

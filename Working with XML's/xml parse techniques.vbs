
'Recurse working folder
WorkDir="C:\ProgramData\ChronoScan\WorkDir\aqilla@F34D8621-ECBA-46F4-9B85-BE6B24C8A6EE"

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
            msgbox oFile.Path
            Set xmlDoc = _
            CreateObject("Microsoft.XMLDOM")

            xmlDoc.Async = "False"
            xmlDoc.Load(oFile.Path)
                
            'xmlDoc.Load("C:\ProgramData\ChronoScan\WorkDir\aqilla@F34D8621-ECBA-46F4-9B85-BE6B24C8A6EE\BatchName_1\batch_config.xml")

            Set Root=xmlDoc.documentElement
            'Set NodeList = xmlDoc.getElementsByTagName("batch_config")

            'Set ElemList = xmlDoc.getElementsByTagName("batch_isprocessed")
            'Set ElemList = xmlDoc.getElementsByTagName("batch_num_errors")
            'Set ElemList = xmlDoc.getElementsByTagName("batch_exported")
            'Set ElemList = xmlDoc.getElementsByTagName("batch_name")

                 'msgbox ElemList.Length
                  
               'For i=0 To (ElemList.length -1)
                   'MsgBox ElemList.item(0).text
               'Next

            set nodes = xmlDoc.selectNodes("//*")    
            For i = 0 to nodes.length-1
               msgbox(nodes(i).nodeName & " - " & nodes(i).text)
            Next

        End If    
        Next 
    End If
End Sub

Function IsAccessible(oFolder)
  On Error Resume Next
  IsAccessible = (oFolder.SubFolders.Count >= 0)
End Function

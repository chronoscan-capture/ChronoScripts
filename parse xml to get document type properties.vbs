'Recurse Type Folder Set your types folder path below
TypeDir="C:\ProgramData\Chronoscan.SAMPLES ONLY\Jobs\Sample invoices full read from PDF@43C51130-9BB4-4EF9-A5EC-3617065F6973\Types"

Set oFSO   = CreateObject("Scripting.FileSystemObject") 
sFileName  = "batch_config.xml"

Set oFolder = oFSO.GetFolder(TypeDir)
Recurse(oFolder)

Sub Recurse(oFldr)
    If IsAccessible(oFolder) Then
        For Each oSubFolder In oFldr.SubFolders
             Recurse oSubFolder
        Next 

        For Each oFile In oFldr.Files
            If Right(oFile.Name,9)=".type.xml" Then 
            'msgbox oFile.Path
            Set xmlDoc = _
            CreateObject("Microsoft.XMLDOM")

            xmlDoc.Async = "False"
            xmlDoc.Load(oFile.Path)
                
            Set Root=xmlDoc.documentElement
            Set NodeList = xmlDoc.getElementsByTagName("C_ChronoScan_Shared_Type")

            ' Get elements by name here
            Set ElemList = xmlDoc.getElementsByTagName("id")  
            Set ElemList1 = xmlDoc.getElementsByTagName("name")            
            Set ElemList2 = xmlDoc.getElementsByTagName("description")
                 'msgbox "Element length " & ElemList.Length
                  
               For i=0 To (ElemList.length -1)
                   'MsgBox ElemList.item(0).text
                   'MsgBox ElemList1.item(0).text
                   'MsgBox ElemList2.item(0).text

                    ' If name matches document type field then get description
                    If ElemList1.item(0).text=UserField_Document_Type.Value Then
                        UserField_Type_Description.value=ElemList2.item(0).text
                    End If
               Next
'/////////////// GET INFO BY LOOPING ALL NODES ////////////////////////////
           'set nodes = xmlDoc.selectNodes("//*")    
            'For i = 0 to nodes.length-1
            'Msgbox (nodes(i).text)
            '    If InStr(nodes(i).text,"Description") Then
                    'msgbox(nodes(i).nodeName & " - " & nodes(i).text)
            '    End If
            'Next
'///////////////////////////////////////////////////////////////////
        End If    
        Next 
    End If
End Sub

Function IsAccessible(oFolder)
  On Error Resume Next
  IsAccessible = (oFolder.SubFolders.Count >= 0)
End Function

' Delete Pages/Document if Image Height < 100 pixels
Set Batch=ChronoApp.GetCurrentBatch

DocNum=ChronoDocument.GetDocNumber
PageCount=ChronoDocument.get_page_count

For i=1 to PageCount
 
    If PageCount > 1 Then
        Height=ChronoDocument.get_page_field_value(i,"ImageHeight")
        If CInt(Height) < 100 Then
            ChronoDocument.DeletePage(i)
            i=i+1
        End If
    Else
        Height=ChronoDocument.get_page_field_value(i,"ImageHeight")
        If CInt(Height) < 100 Then
            Batch.DeleteDocument(DocNum)
        End If
    End If
	
PageCount=ChronoDocument.get_page_count

Next

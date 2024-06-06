Set Batch = ChronoApp.GetCurrentBatch
docCount = Batch.GetDocCount

For i = 0 To docCount-1
    Set Document = Batch.GetDocument(i)
    fieldCount = Document.GetFieldCount
    For f = 0 To fieldCount-1 
        Set field=Document.GetFieldIdx(f)
        field.ValidateStatus = true
    Next    
Next    
Document.Validate

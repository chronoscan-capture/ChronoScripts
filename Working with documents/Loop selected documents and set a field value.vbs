' Get selection
sel = ChronoApp.Indexer_GetSelectedRows

Set Batch = ChronoApp.GetCurrentBatch 

' Loop selection
for i=0 to UBound(sel)
    ' Get document i from the selection
    Set Doc=Batch.GetDocument(sel(i))
    ' Set Doc field NOTES value to position on selection
    Call Doc.set_field_value("NOTES", i+1 & ".sel")
    Next

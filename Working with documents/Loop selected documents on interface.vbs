

Dim arrRows
arrRows = ChronoApp.Indexer_GetSelectedRows()
If UBound(arrRows) > -1 Then
  Set Batch = ChronoApp.GetCurrentBatch()
  For i = 0 To UBound(arrRows)
    Set document=Batch.GetDocument(i+1)             
    msgbox document.get_field_value("Supplier")
  Next
End If

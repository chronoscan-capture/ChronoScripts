Set Batch = ChronoApp.GetCurrentBatch ' Set batch to currently opened batch object
If Not Batch Is Nothing Then
  Set Doc = Batch.getDocument(0)
  If Not Doc Is Nothing Then
    Doc.ignore(1)
    isignored = Doc.isIgnored()
    MsgBox "Document " & isignored
  End If
End If

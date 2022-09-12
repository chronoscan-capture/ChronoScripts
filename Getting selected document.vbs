' Returns first selected document from selection (if any)
Set Batch = ChronoApp.GetCurrentBatch ' Set batch to currently opened batch object
Set Doc = Batch.GetSelectedDocument()

fieldValue = Doc.GetDocNumber()
Call Doc.IgnorePage(1, 1)' for example; we ignore first page of selected doc

msgbox "Document " & (fieldValue+1) & " is selected"

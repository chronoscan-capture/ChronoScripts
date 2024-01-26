Set batch = ChronoApp.GetCurrentBatch()
numDocs=batch.GetDocCount()
For doc = 0 To numDocs-1
    Set docObj=batch.GetDocument(doc)
    msgbox docObj.get_field_value("expinfo.XML Report Extended.File System xml output") ' ChronoScan saves the generated file after exporto on a field called: "expinfo.converter module.outputmodule"
Next

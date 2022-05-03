Set Batch = ChronoApp.GetCurrentBatch ' Set batch to currently opened batch object

Dim NumDocs
NumDocs = Batch.GetDocCount ' Get the number of documents in the batch

' ignore **first** document
Call Batch.IgnoreDocument(1, 1)

' ignore **last** document
Call Batch.IgnoreDocument(NumDocs, 1)


' Loop the batch and tell me which documents are ignored
For numDoc = 0 To NumDocs-1

    ' Ask if document is ignored
    Dim isIgnored
    isIgnored = Batch.IsDocumentIgnored(numDoc+1)
    
    If isIgnored = 1 Then
        msgbox "Document " & numDoc+1 & " is IGNORED" 
    End If

Next

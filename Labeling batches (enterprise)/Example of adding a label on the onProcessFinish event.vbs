
' This is a basic example of how to label everybatch that is successfully processed with a ChatGPT request

' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'   * Only enterprise Jobs
'   * The "CHATGPT" label must exists in the working Job 
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

' Place this script in the "OnProcessFinish" event:

Set Batch = ChronoApp.GetCurrentBatch ' Set batch to currently opened batch object

Dim NumDocs
Dim NumRequestOk
NumRequestOk = 0
NumDocs=Batch.GetDocCount ' Get the number of documents in the batch
 
' Loop the batch
For numDoc = 0 To NumDocs-1

    ' Get the document object
    Set Doc=Batch.GetDocument(numDoc)
    'Get a field value for current document
    fieldvalue = Doc.get_field_value("sysval_oai")
    
    If fieldvalue = "1" Then
        NumRequestOk = NumRequestOk + 1
    End If

Next

If NumDocs = NumRequestOk Then
    ' add existing label for the job
    res = Batch.AddLabel("CHATGPT")
Else
    ' remove existing label for the job
    res = Batch.RemoveLabel("CHATGPT")
End If

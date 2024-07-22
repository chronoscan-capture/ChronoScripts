
' This is a basic example of how to label everybatch that is successfully processed with a ChatGPT request

' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'      The "CHATGPT" label must exists in the working Job 
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

' Place this script on the "OnDocumentProcessFinish" document event

Dim val
val = ChronoDocument.get_field_value("sysval_oai")

' get the current batch
Set ChronoBatch = ChronoApp.GetCurrentBatch()
If val = "1" Then
    ' add and existing label for the job
    res = ChronoBatch.AddLabel("CHATGPT")
Else
    ' add and existing label for the job
    res = ChronoBatch.RemoveLabel("CHATGPT")
End If


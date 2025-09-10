
' Adding a label to a batch
Dim res

' get the current batch
Set batch = ChronoApp.GetCurrentBatch()

' add and existing label for the job
res = batch.AddLabel("DUPLICATE")

If res = 1 Then
    MsgBox "Label added"
Else
    MsgBox "Label could not be added"
End If    

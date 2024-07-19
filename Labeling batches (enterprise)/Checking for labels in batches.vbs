' Adding a label to a batch
Dim res
Dim label
label = "DUPLICATE"

' get the current batch
Set ChronoBatch = ChronoApp.GetCurrentBatch()

' add and existing label for the job
res = ChronoBatch.HasLabel(label)

If res = 1 Then
    MsgBox "Batch has the label " & label
Else
    MsgBox "Batch does not have the label " & label
End If    

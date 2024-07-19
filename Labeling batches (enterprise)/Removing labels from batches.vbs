
' Adding a label to a batch
Dim res

' get the current batch
Set ChronoBatch = ChronoApp.GetCurrentBatch()

' add and existing label for the job
res = ChronoBatch.RemoveLabel("DUPLICATE")

If res = 1 Then
    MsgBox "Label removed"
Else
    MsgBox "Label not removed"
End If    

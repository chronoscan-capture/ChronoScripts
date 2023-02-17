' Since ChronoScan Version v1.0.2.96

Dim res

' Getting a document enterprise status
res = ChronoDocument.GetEntStatus()

ChronoApp.AddToOutputWindow("Status is " & res)

' Setting a document status
' accepted status;
'   toprocess, toreprocess, topreclass, topostclass, toreview, toconfigure, manual_index, exported, 
'   toapprove, waiting_approval, rejected, approved, system_indexed, user_indexed, processed
res = ChronoDocument.SetEntStatus("toreview")

If res = 1 Then
    res = "Success"
Else
    res = "Error"
End If

ChronoApp.AddToOutputWindow("Response is " & res)


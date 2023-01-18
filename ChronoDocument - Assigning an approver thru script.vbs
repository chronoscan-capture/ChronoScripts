
' Get the current job
Set job = ChronoApp.GetCurrentJob()

' msgbox job.JobName

Dim username
username = "JApprover"

' get an user Id
Dim approverid
approverid = ChronoApp.GetUserId(username)

' msgbox " user " & username & " has id " & approverid

' check if the job ha approval workflow active
If job.IsApprovalWorkflow Then

    ' check if user is approver
    If job.IsUserApprover(username) Then
        
        ' assign the user as the approver for (current document)
        Dim assigned
        assigned = ChronoDocument.SetApprover(approverid)
        
        If assigned Then
            msgbox "The user " & username & " has been assigned to approve the current document"
        End If

    End If

End If

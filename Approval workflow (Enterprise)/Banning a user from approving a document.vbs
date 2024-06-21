
' This script is placed on the enterprise event: OnENT_ValidateDocument
' thit event is triggered when an user clicks the validate document button on the indexer


' validate the document first, cancel following operations if not valid (optional)
Dim valid
Call ChronoDocument.Validate()
valid = ChronoDocument.GetValidateStatus()

If valid = 1 Then
    ' 1. check if job has approval workflow   
    Set job = ChronoApp.GetCurrentJob()
    Dim isApproval
    isApproval = job.IsApprovalWorkflow()
    If isApproval = 1 Then 
        ' we probably only want to ban the user if we are indexing/validating the document
        Dim stage ' 0 configurating stage, 1  index mode, 2 approval stage, -1 undefined
        stage = ChronoApp.Indexer_GetEntStage()
        If stage = 1 Then
            ' get the current user
            Dim current_userId
            current_userId = ChronoApp.GetCurrentUserId()
        
            ' Ban the user from approving this doc
            Dim result
            result = ChronoDocument.banApprover(current_userId)

            If result = 1 Then
                msgbox "user " & ChronoApp.GetUserName(current_userId) & " won't be able to approve this document"
            Else 
               ' msgbox  "..."
            End If
        
        End If
    End If
Else 
    ChronoProc.add_retvalue("**CANCEL**") 'return this to cancel enterprise indexer validation behaviour
End If
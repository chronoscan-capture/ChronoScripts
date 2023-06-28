' Useful script to loop a batch of documents
Set Batch = ChronoApp.GetCurrentBatch ' Set batch to currently opened batch object

Dim NumDocs

NumDocs=Batch.GetDocCount ' Get the number of documents in the batch

'counters
Dim correct 
correct = 0
Dim wrong
wrong = 0
 
' Loop the batch
For numDoc = 0 To NumDocs-1

    ' Get the document object
    Set Doc=Batch.GetDocument(numDoc)
    'Get a field value for current document
    TRUE_CLASS = Doc.get_field_value("True_class")
    PRED_CLASS = Doc.get_field_value("Predicted_Class")
    
    if TRUE_CLASS = PRED_CLASS Then
        correct = correct + 1
    else 
        wrong = wrong + 1
    End if

Next

Dim correct_pct
correct_pct = (correct / NumDocs) * 100
Dim wrong_ptc 
wrong_pct = (wrong / NumDocs) * 100

msgbox "total docs: " & NumDocs & ". Pct correct: " & correct_pct & "%, " & " pct wrong: " & wrong_pct & "%"
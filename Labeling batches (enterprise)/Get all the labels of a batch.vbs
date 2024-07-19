' Get the current batch
Set ChronoBatch = ChronoApp.GetCurrentBatch()

' Declare variables
Dim all_labels
Dim labels_array

' Initialize all_labels as an empty string
all_labels = ""

' Get the array of labels
labels_array = ChronoBatch.GetLabels()

' Iterate through the array of labels and concatenate them
For Each strLabel In labels_array
    ' If all_labels is not empty, add the delimiter
    If all_labels <> "" Then
        all_labels = all_labels & " | "
    End If
    all_labels = all_labels & strLabel
Next

' Display the result in a message box
MsgBox all_labels

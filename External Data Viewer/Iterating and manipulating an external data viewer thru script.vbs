' This sample uses methods of the ChronoExtDBViewer object
' Methods shown here are available in ChronoScan since v.1.0.2.95

Dim ret
Dim row
Dim col
Dim column_type
Dim column_name
Dim numeric_val
Dim string_val
Dim compare_val
compare_val = 5.5
Dim string_to_find 
string_to_find = "JOHN DOE"
Dim err_log

Dim TOTAL_ROWS
TOTAL_ROWS = ChronoExtDBViewer.GetRowCount()
Dim TOTAL_COLS 
TOTAL_COLS = ChronoExtDBViewer.GetColCount()

' iterating data viewer results table
For row = 0 To TOTAL_ROWS-1 

    ' logging output for each row
    ' ChronoApp.AddToOutputWindow("Iterating Row " & row)

    For col = 0 To TOTAL_COLS-1

        ' getting the column name
        column_name = ChronoExtDBViewer.GetColName(col)
        ' getting every column type
        column_type = ChronoExtDBViewer.GetColType(col)
        ' logging output for each column
        ' ChronoApp.AddToOutputWindow("Iterating Row " & row & ", column " & col & ", column name: " & column_name & ", column type:" & column_type)

        If column_type = "numeric" Then

            ' doing something with a numeric value
            numeric_val = ChronoExtDBViewer.GetCellValueNumeric(row, col)
            If numeric_val < compare_val Then
                ' if cell value lower than compare value we set a red background to the cell, for example
                ret = ChronoExtDBViewer.AddStyleClass(row, col, "danger")
            End If

        Else

            ' doing something with a string value
            string_val = ChronoExtDBViewer.GetCellValueString(row, col)
            If string_val = string_to_find Then
                ' if cell value contains the string we are looking for, we set a green background, for example
                ret = ChronoExtDBViewer.AddStyleClass(row, col, "success")
            End If

        End If

    Next 

Next
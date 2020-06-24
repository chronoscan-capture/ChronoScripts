numRows = ChronoApp.GlobalListSize("list_1")


For numRow = 0 To numRows-1
    msg="row: "&numRow & ", "
    myArray = ChronoApp.GlobalListGet("list_1",numRow)
    If UBound(myArray) > 0 Then
        msg=msg&" size:" & UBound(myArray) & "   "
        For i = 0 To UBound(myArray)
            msg=msg & "|" & myArray(i)
        Next
        MsgBox msg
    End If
Next

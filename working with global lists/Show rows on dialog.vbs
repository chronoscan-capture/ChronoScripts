Call ChronoApp.dlg_seloption_prepare("List sample", "description list sample")

'add my 5 columns
Call ChronoApp.dlg_seloption_add_column("String field 1", "sz:80|align:Left" )
Call ChronoApp.dlg_seloption_add_column("Date field", "sz:120|align:Center" )
Call ChronoApp.dlg_seloption_add_column("inteter value", "sz:80|align:Right" )
Call ChronoApp.dlg_seloption_add_column("double value", "sz:80|align:Right" )
Call ChronoApp.dlg_seloption_add_column("String field 2", "sz:80|align:Left" )


numRows = ChronoApp.GlobalListSize("list_1")
For numRow = 0 To numRows-1
    myArray = ChronoApp.GlobalListGet("list_1",numRow)

    If UBound(myArray) > 0 Then
        idx = ChronoApp.dlg_seloption_add_row(numRow)
        For i = 0 To UBound(myArray)
            Call ChronoApp.dlg_seloption_setrowtext(idx, i+1, myArray(i))
        Next
    End If
Next

rowselected = ChronoApp.dlg_seloption_show()

MsgBox "Row selected: " & rowselected

' external data viewer update example
' a working external data view must be running with a corresponding UPDATE statement configured

' This example would execute the following configured statement for an example view "test_view"

        ' UPDATE [test_view]
		'  SET Activity = [INPARAM_1] 
		'     WHERE 
		' Supplier = [INPARAM_2] AND Invoice_number = [INPARAM_3] AND LineId = [INPARAM_4]

' result
Dim r 

Dim newValue
newValue = "New value"

' get the lineId from a selected row from the view, for example
Dim LineId
LineId = ChronoExtDBViewer.CurrentSelectedGetFieldValue("LineId")

If LineId <> "" Then
    ' Params: "test view"(View name), Activity[INPARAM_1], Supplier[INPARAM_2], Invoice_number[INPARAM_3], LineId[INPARAM_4]
    r =  ChronoExtDBViewer.ExecuteUpdate("test_view", newValue, UserField_Supplier.value, UserField_Invoice_Number.value, CInt(LineId))
    If r > 0 Then
        msgbox "Update success"
    Else 
        msgbox "Error while updating"
    End If
Else 
    msgbox "No row selected, please select one"
End If
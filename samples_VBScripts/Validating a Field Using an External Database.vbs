' https://www.chronoscan.org/doc/validating_a_field_using_an_external_database.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'Add this script to the OnValueChanged property:
 
 
Set MyDB = ChronoApp.GetChronoScanDBConnection("MyDatabase", "", "")
 
' Build your SQL Query
SQLString = "Select * from Suppliers where AccountRef = '" & UserField_Document_Type.value & "'"
 
Set rsCustomers = MyDB.Execute(SQLString)
 
If Not rsCustomers.EOF Then
    UserField_Document_Type.ValidateStatus = 1
Else
    UserField_Document_Type.ValidateStatus = 0   
End If
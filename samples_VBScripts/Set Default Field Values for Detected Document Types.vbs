'https://www.chronoscan.org/doc/set_default_field_values_for_detected_document_types.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'This sample script allows you to set a default value for your fields when a Document Type is assigned or detected.
 
 
Dim SQLString
 
' Connect to the database
Set MyDB = ChronoApp.CreateAdoDBConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Dbs\suppliers.mdb;Persist Security Info=False;","","")
 
SQLString = "Select * from Suppliers where AccountRef = '" & UserField_Document_Type.value & "'"
 
' Search a value based on the document type
Set rsCustomers = MyDB.Execute(SQLString)
 
If Not rsCustomers.EOF Then
UserField_AccountRef.value = rsCustomers.Fields("NominalCode")
' This will set a new default value for the field AccountRef for the current Document Type.
Call DocumentType.SetFieldDefaultValueForType("AccountRef",UserField_AccountRef.value)
End If
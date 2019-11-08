' https://www.chronoscan.org/doc/lookup_fields_from_an_external_database_using_scripts.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

Dim SQLString
 
' Create a database connection,
Set MyDB = ChronoApp.CreateAdoDBConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Dbs\suppliers.mdb;Persist Security Info=False;", "", "")
 
' Build your SQL Query
SQLString = "Select * from Suppliers where SupplierName = '" & UserField_Document_Type.value & "'"
 
Set rsCustomers = MyDB.Execute(SQLString)
 
If Not rsCustomers.EOF Then
' Action when the searched value exist
    UserField_AccountRef.value = rsCustomers.Fields("AccountRef")
    UserField_NominalCode.value = rsCustomers.Fields("NominalCode")
    UserField_TaxCode.value = rsCustomers.Fields("TaxCode")
 
' Validate the Fields
    UserField_Document_Type.ValidateStatus = 1
    UserField_AccountRef.ValidateStatus = 1
Else
' Set field validation to error
    UserField_Document_Type.ValidateStatus = 0
    UserField_AccountRef.ValidateStatus = 0
    UserField_AccountRef.ValidateMessage = "Supplier doesn't exist On the remote database"
End If



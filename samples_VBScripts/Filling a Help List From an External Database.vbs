' https://www.chronoscan.org/doc/filling_a_help_list_from_an_external_database.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'This sample script shows how to fill a helplist on ChronoScan using an external database:
 
'Example 1:
 
Dim SQLString
 
Set MyDB = ChronoApp.CreateAdoDBConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\my_dbs\suppliers.mdb;Persist Security Info=False;", "", "")
 
SQLString = "Select * from Suppliers Order By SupplierName"
 
Set rsCustomers = MyDB.Execute(SQLString)
 
Call UserField_Document_Type.HelpList_Clear()
 
Do While Not rsCustomers.EOF
    Call UserField_Document_Type.HelpList_AddValue( rsCustomers.Fields("SupplierName") )
 
    rsCustomers.MoveNext
Loop
 
Call UserField_Document_Type.HelpList_Populate()
 
 
 
'Example 2:
 
UserField_Invoice_Number.HelpList_Clear()
 
Call UserField_Invoice_Number.HelpList_AddValueDescription( "Value 1", "Description 1" )
Call UserField_Invoice_Number.HelpList_AddValueDescription( "Value 2", "Description 2" )
Call UserField_Invoice_Number.HelpList_AddValueDescription( "Value 3", "Description 3" )
Call UserField_Invoice_Number.HelpList_AddValueDescription( "Value 4", "Description 4" )
 
UserField_Invoice_Number.HelpList_Populate()
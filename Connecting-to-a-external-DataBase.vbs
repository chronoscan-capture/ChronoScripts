
'For ChronoScan 1.0.0.49 or greater.
'We have added a new function to ChronoScan VBscritps that allows to maintain an open database connection, this will speed up queries to external data sources.

'The new function is:

'ChronoApp.CreateAdoDBConnection( DSN, USERNAME, PASSWORD)

'DSN: your connection String
'USERNAME: User name for the connection, (if not supplied on the DSN)
'PASSWORD: Password for the connection, (if not supplied on the DSN)
'To create a new persistent database connection replaces the old ADODB.Connection method:

Set MyDB = CreateObject("ADODB.Connection")

MyDB.Open "ULISES_JOSE_DEV"

'With a call to the new function:

Set MyDB = ChronoApp.CreateAdoDBConnection("ULISES_JOSE_DEV", "", "")

'Very Important:
'Remember to remove all MyDB.Close calls, now ChronoScan will manage the database flow to speed up calls. Avoid this steep may stop your connection to work.

' Useful script to loop a batch of documents

Set Batch = ChronoApp.GetCurrentBatch ' Set batch to currently opened batch object

Dim NumDocs

NumDocs=Batch.GetDocCount ' Get the number of documents in the batch
 
' Loop the batch
For numDoc = 0 To NumDocs-1

	' Get the document object
	Set Doc=Batch.GetDocument(numDoc)
	'Get a field value for current document
	fieldvalue = Doc.get_field_value( "field name" )

Next


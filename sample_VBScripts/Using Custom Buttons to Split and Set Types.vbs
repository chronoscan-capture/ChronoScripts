' https://www.chronoscan.org/doc/using_custom_buttons_to_split_and_set_types.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'Using the Custom Buttons functionality it is possible to set the type for the selected document and split on that same document. The code needed is as follows and it should go on the OnClick section:
 
 
'Get selected page number
 
selectedPage = ChronoDocument.GetDocFirstSelectedPage
 
'Get current document number and the next document in line
  
currDocNum=ChronoDocument.GetDocNumber()
nexDocNum=currDocNum+1
 
'This will split the document on the desired page
 
ChronoDocument.SplitOnPage(SelectedPage)
 
Set Batch=ChronoApp.GetCurrentBatch()
 
'This part of the code will set the type for the next document after the split
 
If nexDocNum < Batch.GetDocCount() Then
       
Set NextDoc=Batch.GetDocument(nexDocNum)
 
NextDoc.SetDocumentType("Name of the type")
 
End If


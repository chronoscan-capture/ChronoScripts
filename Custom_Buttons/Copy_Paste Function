' Copy script for selecting page for copy/paste function, uncomment msgbox if needed for debug
' Add script to a custom "copy" button

'Get current batch object
Set Batch=ChronoApp.GetCurrentBatch
' Get selected document number
numDoc=ChronoDocument.GetDocNumber
' Get selected page
selPage=ChronoDocument.GetDocFirstSelectedPage
'msgbox selPage
' Get the path and file for the selected page
pageFile=ChronoDocument.get_doc_page_file(selPage+1)
' Save the selected file to a global variable
Call ChronoApp.SetGlobalVariable("fileTopaste",pageFile)
'msgbox pageFile

----------------------

' Paste script for selecting target page for copy/paste function, uncomment msgbox if needed for debug
' Add script to a custom "paste" button

' Get selected page to copy before
selPage=ChronoDocument.GetDocFirstSelectedPage
'Get page to paste from global variable
pageTopaste=ChronoApp.GetGlobalVariable("fileTopaste",DefaultValue)
' Paste page
Call ChronoDocument.AddPage(pageTopaste,selPage)

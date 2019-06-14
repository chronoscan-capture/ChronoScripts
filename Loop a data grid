' Useful script to loop a data grid and get or set values
Dim numRows
Dim GridPanel
 
GridPanel=1'This is the panel number for the first panel
numRows=ChronoDocument.GetXgridRowCount (GridPanel) ' Retrieve the number of rows in the datagrid

' Loop your grid rows here
Dim row

For row = 0 To numRows-1

    YESANS=ChronoDocument.GetXgridFieldValue (GridPanel, row, "YES") ' Get grid row values by column name here
    NOANS=ChronoDocument.SetXgridFieldValue (GridPanel, row, "NO","NEW VALUE")   ' Set grid row values by column name here


Next

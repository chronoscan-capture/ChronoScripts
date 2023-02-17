'https://www.chronoscan.org/doc/sum_a_capture_grid_column_.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'This example shows how to loop capture grid records and extract the data from a column to sum the values to compare it to a "Net Total" field

SetLocale("en-us")
 
Dim numRows
Dim GridPanel
 
GridPanel=1'first panel
numRows=ChronoDocument.GetXgridRowCount (GridPanel)
 
Dim Total
Total=0
Dim row
For row = 0 To numRows-1
Dim lineTotal
    lineTotal=ChronoDocument.GetXgridFieldValue (GridPanel, row, "Line total")
    Total=Total + CDbl(lineTotal)
Next
 
 
' Validate values
dif = Abs(CDbl(UserField_Net_Total.value) - Total)
If dif <= 0.03 Then
    UserField_Net_Total.ValidateStatus = true
Else
    UserField_Net_Total.ValidateStatus = false
    UserField_Net_Total.ValidateMessage = "Net Total is not equal to Sum all lines " & Total
End If
 
UserField_Net_Total.Validate
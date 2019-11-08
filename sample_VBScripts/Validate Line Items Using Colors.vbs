' https://www.chronoscan.org/doc/validate_line_items_using_colors.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'It is possible to create validation rules for line items and color the cell where the error is present to make it easier to visually identify them. A video with more details can be found here.
 
 
 
' This technique will require modification for different types of grids. The code on this sample assumes you have a quantity and a price column on your grid and it uses these two columns to check if the total column is correctly calculated.
' This sample also uses a Grid Total field on the Form View that will display the calculated total from the grid using the Line Total column. That same Grid Total field will also display a validation error if the Grid Validation script fails.
' This is the code that will perform the grid validation and calculate the value for the Grid Total field, this code goes into OnDocumentProcessFinish:
 
SetLocale("en-US")
'Set Modules Panel Grid Control 1
PanelNo = 1
Total = CDbl(0)
 
'Get total number of rows read for Grid
NumRows = ChronoDocument.GetXgridRowCount(PanelNo)
 
'Loop the grid
For i = 0 To NumRows - 1
    'Check to make sure empty values don't cause problems
    If ChronoDocument.GetXgridFieldValue (PanelNo,i,"Quantity") = "" Then
        Qty = CDbl(0)
    Else
        Qty = CDbl(ChronoDocument.GetXgridFieldValue (PanelNo,i,"Quantity"))
    End If
   
    'Check to make sure empty values don't cause problems
    If ChronoDocument.GetXgridFieldValue (PanelNo,i,"Unit Price") = "" Then
        UPrice = CDbl(0)
    Else
        UPrice = CDbl(ChronoDocument.GetXgridFieldValue (PanelNo,i,"Unit Price"))
    End If
   
    'Check to make sure empty values don't cause problems
    If ChronoDocument.GetXgridFieldValue (PanelNo,i,"Line total") = "" Then
        CurrLine = CDbl(0)
    Else
        CurrLine = CDbl(ChronoDocument.GetXgridFieldValue (PanelNo,i,"Line total"))
    End If
   
    'Calculate line total from Quantity and Price
    LineTotal = FormatNumber(Qty * UPrice)
   
    'Set calculated totals to CALC TOTAL column
    Call ChronoDocument.SetXgridFieldValue (PanelNo,i,"CALC TOTAL", LineTotal)
    Call ChronoDocument.SetXgridFieldValue (PanelNo,i,"Dif",LineTotal - CurrLine)
 
    'Check calculated total against captured total and color cell when validation fails
    If LineTotal <> FormatNumber(CurrLine) Then
        Call ChronoDocument.SetXgridFieldColor(PanelNo,i,"Line total",rgb(255,0,0))
    Else
        Call ChronoDocument.SetXgridFieldColor(PanelNo,i,"Line total",-1)
    End If   
 
    Total = CDbl(Total + CurrLine)
 
Next
 
'msgbox  FormatNumber(ChronoDocument.GetXgridFieldValue (PanelNo,18,"Quantity") * ChronoDocument.GetXgridFieldValue (PanelNo,18,"Unit Price"))
 
'msgbox FormatNumber(Total)
'Set calculated grid total to the Form View Grid Total field
UserField_Grid_Total.value = FormatNumber(Total)
 
 
Additional validation can be added into OnValidate:
 
On Error Resume Next
SetLocale("en-US")
' Set Error by default
UserField_Invoice_Total.ValidateMessage = "Error: Net Total+Tax Total is not equal to Invoice Total"
UserField_Invoice_Total.ValidateStatus = false
 
' Get Values
 Subtotal = CDbl(UserField_Net_Total.value)
tax = CDbl(UserField_Tax_Total.value)
total = CDbl(UserField_Invoice_Total.value)
transport = CDbl(UserField_Transport_Total.value)
' Validate values
dif = Abs(total - (subtotal + tax + transport))
If dif <= 0.03 Then
    UserField_Invoice_Total.ValidateStatus = true
End If
If UserField_Invoice_Total.value = "" Then
    UserField_Invoice_Total.ValidateStatus = false
End If
 
If CDbl(UserField_Net_Total.value) + CDbl(UserField_Transport_Total.value) = CDbl(UserField_Grid_Total.value) Then
    UserField_Grid_Total.ValidateStatus = 1
Else
    UserField_Grid_Total.ValidateStatus = 0
    UserField_Grid_Total.ValidateMessage = "There is something wrong with the line items"
End If
 
Also, extra functionality can be added by using custom buttons. This code will help when there are missing Line Total values on the grid. The button will fill in the calculated value when the Line Total value is missing.
This code goes inside OnButtonClick for Button 1:
 
SetLocale("en-US")
PanelNo = 1
Total = CDbl(0)
 
NumRows = ChronoDocument.GetXgridRowCount(PanelNo)
 
For i = 0 To NumRows - 1
  
    If ChronoDocument.GetXgridFieldValue (PanelNo,i,"Quantity") = "" Then
        Qty = CDbl(0)
    Else
        Qty = CDbl(ChronoDocument.GetXgridFieldValue (PanelNo,i,"Quantity"))
    End If 
 
   
    If ChronoDocument.GetXgridFieldValue (PanelNo,i,"Unit Price") = "" Then
        UPrice = CDbl(0)
    Else
        UPrice = CDbl(ChronoDocument.GetXgridFieldValue (PanelNo,i,"Unit Price"))
    End If   
 
   
    If ChronoDocument.GetXgridFieldValue (PanelNo,i,"Line total") = "" Then
        CurrLine = CDbl(0)
    Else
        CurrLine = ChronoDocument.GetXgridFieldValue (PanelNo,i,"Line total")
    End If
   
    LineTotal = FormatNumber(Qty * UPrice)
 
    If CurrLine = 0 Then
        Call ChronoDocument.SetXgridFieldValue(PanelNo,i,"Line total",LineTotal)
        CurrLine = LineTotal
    End If 
 
    If LineTotal <> FormatNumber(CurrLine) Then
        Call ChronoDocument.SetXgridFieldColor(PanelNo,i,"Line total",rgb(255,0,0))
    Else
        Call ChronoDocument.SetXgridFieldColor(PanelNo,i,"Line total",-1)
    End If   
 
    Total = CDbl(Total + CurrLine)
 
Next
 
'msgbox  FormatNumber(ChronoDocument.GetXgridFieldValue (PanelNo,6,"Quantity") * ChronoDocument.GetXgridFieldValue (PanelNo,6,"Unit Price"))
 
'msgbox FormatNumber(Total)
UserField_Grid_Total.value = FormatNumber(Total)
 


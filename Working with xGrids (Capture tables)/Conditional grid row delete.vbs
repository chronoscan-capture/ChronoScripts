' Grid Cleanup
Dim GridPanel2
GridPanel2=2
numRows=ChronoDocument.GetXgridRowCount(GridPanel2)


dim row
row=0

Do While row < numRows
    ACT=ChronoDocument.GetXgridFieldValue(GridPanel2,row,"ALLOWANCE_CHARGE_TOTAL")
    'msgbox ACT & " " & " " & numRows & " True/false"  & IsBlank(ACT)
    If ACT="" OR ACT="0.00" Then
        Call ChronoDocument.DeleteXgridRow(GridPanel2,row)
        numRows=numRows-1
    Else
        row=row+1
    End If
Loop

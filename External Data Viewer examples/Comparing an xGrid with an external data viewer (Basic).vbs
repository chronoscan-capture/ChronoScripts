' Basic example comparing an xgrid table with an external database viewer in matching mode
' see https://chronoscan.org/doc/chronoextdbviewer.htm for full ChronoExtDBViewer object

Set Batch = ChronoApp.GetCurrentBatch
Set curDoc = Batch.getSelectedDocument

Dim tableNumber, numRows
tableNumber=1'we are using table 1
 
numRows=Batch.GetXgridRowCount (curDoc.GetDocNumber , tableNumber)
' MsgBox "Total Records in xgrid table: "  + CStr(numRows)            

Dim lineMismatches
lineMismatches = 0
Dim anyMismatch
Dim linesId

'navigating xgrid table records
Dim row
For row = 0 To numRows-1
 
        anyMismatch = 0

       ' xgrid manipulation
        Dim xgrid_lineId 
        Dim xgrid_Service
        Dim xgrid_Activity
        Dim xgrid_Quantity
        Dim xgrid_Rate
        Dim xgrid_Amount

        ' line id in this case is a chronoScan auto-generated value
        xgrid_lineId = "000" & row+1
        ' get the info from xgrid we want to compare with the external database        
        xgrid_Service = Batch.GetXgridFieldValue (numDoc, tableNumber, row, "Service")
        xgrid_Activity = Batch.GetXgridFieldValue (numDoc, tableNumber, row, "Activity")
        xgrid_Quantity = Batch.GetXgridFieldValue (numDoc, tableNumber, row, "Quantity")
        xgrid_Rate = Batch.GetXgridFieldValue (numDoc, tableNumber, row, "Rate")
        xgrid_Amount = Batch.GetXgridFieldValue (numDoc, tableNumber, row, "Amount")        

        ' MsgBox "Xgrid line Id is: " +xgrid_lineId
        
        ' External dbviewer manipulation
        Dim exdbview_lineId 
        Dim exdbview_lineId_colIndex
        Dim exdbview_Quantity
        Dim exdbview_Quantity_colIndex
        Dim exdbview_Rate
        Dim exdbview_Rate_colIndex
        Dim exdbview_Amount
        Dim exdbview_Amount_colIndex
        
        exdbview_lineId_colIndex = ChronoExtDBViewer.GetColIndex("LineId")
        If exdbview_lineId_colIndex <> -1 Then
            
            ' Comparing line id
            exdbview_lineId = ChronoExtDBViewer.GetCellValueNumeric(row, exdbview_lineId_colIndex)
            exdbview_lineId = "000" & exdbview_lineId
            If exdbview_lineId <> xgrid_lineId  Then
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_lineId_colIndex, "danger")
                anyMismatch = anyMismatch+1
            Else 
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_lineId_colIndex, "success")
            End If

            ' Comparing quantity
            exdbview_Quantity_colIndex = ChronoExtDBViewer.GetColIndex("Quantity")
            exdbview_Quantity = ChronoExtDBViewer.GetCellValueNumeric(row, exdbview_Quantity_colIndex)
            If CDbl(exdbview_Quantity) <> CDbl(xgrid_Quantity)  Then
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_Quantity_colIndex, "danger")
                anyMismatch = anyMismatch+1
            Else 
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_Quantity_colIndex, "success")
            End If

            ' Comparing rate
            exdbview_Rate_colIndex = ChronoExtDBViewer.GetColIndex("Rate")
            exdbview_Rate = ChronoExtDBViewer.GetCellValueNumeric(row, exdbview_Rate_colIndex)
            If CDbl(xgrid_Rate) <> CDbl(exdbview_Rate)  Then
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_Rate_colIndex, "danger")
                anyMismatch = anyMismatch+1
            Else 
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_Rate_colIndex, "success")
            End If

            ' Comparing Amount
            exdbview_Amount_colIndex = ChronoExtDBViewer.GetColIndex("Amount")
            exdbview_Amount = ChronoExtDBViewer.GetCellValueNumeric(row, exdbview_Amount_colIndex)
            If CDbl(xgrid_Amount) <> CDbl(exdbview_Amount)  Then
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_Amount_colIndex, "danger")
                anyMismatch = anyMismatch+1
            Else 
                ret = ChronoExtDBViewer.AddStyleClass(row, exdbview_Amount_colIndex, "success")
            End If

        End If

        If anyMismatch > 0 Then
            lineMismatches = lineMismatches + 1
            If linesId <> "" Then
                linesId = linesId & "," & xgrid_lineId 
            Else 
                linesId = xgrid_lineId
            End If
        End If

Next

' summary
If lineMismatches > 0 Then
    msgbox "There is/are " & lineMismatches & " mismatch(es) on this document on line(s) " & linesId
End If
' Since ChronoScan Version v1.0.2.96

Dim res

' Approve a document
res = ChronoDocument.Approve()

' reject a document 
res = ChronoDocument.Reject("i.e: Wrong department")

' send to revise a document 
res = ChronoDocument.SendToRevise("i.e: It Needs totals revision")

' set as waiting approval
res = ChronoDocument.SetEntStatus("waiting_approval")


Attribute VB_Name = "modPaymentMaster"


Public Function GeneratePaymentID() As String

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("Payment_Master")
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    GeneratePaymentID = "PAY" & Format(lastRow, "0000")

End Function

Attribute VB_Name = "modprintformat"
Sub LoadLogo()

    Dim ws As Worksheet
    Dim logoPath As String
    Dim shp As Shape

    Set ws = Sheets("Print_Format")

    logoPath = Sheets("Settings").Range("B11").value & "\logo.jpg"

    On Error Resume Next
    ws.Shapes("CompanyLogo").Delete
    On Error GoTo 0

    If Dir(logoPath) <> "" Then
        Set shp = ws.Shapes.AddPicture(logoPath, msoFalse, msoTrue, 10, 10, 80, 50)
        shp.name = "CompanyLogo"
    Else
        MsgBox "Logo file not found!", vbExclamation
    End If

End Sub

Sub CreateProfessionalReceipt()

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Sheets("Print_Format")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Sheets.Add
        ws.name = "Print_Format"
    End If
    
    ws.Cells.Clear
    
    ' Column width
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 20
    
    ' HEADER
    ws.Range("B2:D2").Merge
    ws.Range("B2").value = "GLOBAL IT SOLUTIONS"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    
    ws.Range("B3:D3").Merge
    ws.Range("B3").value = "Address Here"
    
    ws.Range("B4:D4").Merge
    ws.Range("B4").value = "?? 9777971045 (Call / WhatsApp)"
    
    ' LINE
    ws.Range("B5:D5").Borders(xlEdgeBottom).Weight = xlMedium
    
    ' INVOICE INFO
    ws.Range("B7").value = "Invoice No:"
    ws.Range("C7").value = "AUTO"
    
    ws.Range("D7").value = "Time:"
    
    ws.Range("B8").value = "Service ID:"
    
    ' CUSTOMER
    ws.Range("B10").value = "Customer:"
    ws.Range("C10:D10").Merge
    
    ws.Range("B11").value = "Mobile:"
    ws.Range("B12").value = "Address:"
    ws.Range("C12:D12").Merge
    
    ' LINE
    ws.Range("B13:D13").Borders(xlEdgeBottom).Weight = xlThin
    
    ' PRODUCT
    ws.Range("B15").value = "Product:"
    ws.Range("B16").value = "Model:"
    ws.Range("B17").value = "Serial:"
    ws.Range("B18").value = "Status:"
    ws.Range("B19").value = "Warranty:"
    
    ' IMAGE BOX
    ws.Range("C15:D19").Merge
    ws.Range("C15:D19").BorderAround 1
    
    ' ACCESSORIES HEADER
    ws.Range("B21").value = "Accessories"
    ws.Range("B21").Font.Bold = True
    
    ws.Range("B22").value = "Name"
    ws.Range("C22").value = "Working"
    ws.Range("D22").value = "Serial"
    
    ws.Range("B22:D22").Borders(xlEdgeBottom).Weight = xlThin
    
    ' AMOUNT
    ws.Range("B27").value = "Total:"
    ws.Range("B28").value = "Advance:"
    ws.Range("B29").value = "Due:"
    
    ' VERIFY
    ws.Range("B31").value = "Verified By:"
    
    ' FOOTER
    ws.Range("B34").value = "Thank You – Visit Again"
    
    ws.Range("D34").value = "Signature"
    
    ' REMOVE ALL GRID LOOK
    ws.Cells.Borders.LineStyle = xlNone
    
    MsgBox "Professional Receipt Ready!"

End Sub

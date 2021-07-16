' TO-DO
' -----
' * add conditional formatting to FormatSpreadsheet
' * add comments to code
'


Sub FormatSpreadsheet()
    Dim base_width As Double
    
    Dim sheet_title_range As String
    
    base_width = 8.43
    
    sheet_title_range = Range("A1:ZZ1")
    
    Columns("B").ColumnWidth = base_width * 2
    Columns("D").ColumnWidth = base_width * 2
    Columns("F").ColumnWidth = base_width * 2.5
    Columns("G").ColumnWidth = base_width * 2.5
    Columns("H").ColumnWidth = base_width * 2.5
    
    Range("A1").Value = UCase("ordem")
    Range("B1").Value = UCase("prioridade")
    Range("C1").Value = UCase("linha")
    Range("D1").Value = UCase("operação")
    Range("E1").Value = UCase("ativo")
    Range("F1").Value = UCase("tipo de manutenção")
    Range("G1").Value = UCase("natureza do serviço")
    Range("H1").Value = UCase("tempo estimado")
    
    Range("A1:H1").VerticalAlignment = xlCenter
    Range("A1:H1").HorizontalAlignment = xlCenter
    
    Range("A1").Font.Bold = True
    Range("B1").Font.Bold = True
    Range("C1").Font.Bold = True
    Range("D1").Font.Bold = True
    Range("E1").Font.Bold = True
    Range("F1").Font.Bold = True
    Range("G1").Font.Bold = True
    Range("H1").Font.Bold = True
    
End Sub

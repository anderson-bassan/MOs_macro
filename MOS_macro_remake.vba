' TO-DO
' -----
' * add comments to code
'


Sub FormatSpreadsheet()
    Dim base_width As Double
    
    base_width = 8.43
    
    sheet_title_range = Range("A1:ZZ1")
    
    Debug.Print VarType(sheet_title_range)
    
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
    
    With Worksheets(1).Range("A:ZZ").FormatConditions _
        .Add(xlBlanksCondition)
        With .Borders
            .Color = RGB(255, 255, 255)
        End With
    End With
    
    With Worksheets(1).Range("A1:ZZ1").FormatConditions _
        .Add(xlNoBlanksCondition)
        With .Interior
            .ColorIndex = 1
        End With
        
        With .Font
            .Bold = True
            .ColorIndex = 2
        End With
    End With
    
    With Worksheets(1).Range("A2:ZZ999").FormatConditions _
        .Add(xlNoBlanksCondition)
        With .Borders
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub

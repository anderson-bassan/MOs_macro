' TO-DO
' -----
' * add comments to code
'


Sub FormatSpreadsheet()
    ' declare the ranges
    Dim title_range As String
    Dim non_title_range As String
    Dim empty_cells_range As String
    
    ' set the base width of the columns
    Dim base_width As Double
    
    base_width = 8.43
    
    ' Set the ranges that will be used
    title_range = "A1:H1"
    non_title_range = "A2:H999"
    empty_cells_range = "A:ZZ"
    
    ' change the width of the used columns
    Columns("B").ColumnWidth = base_width * 2
    Columns("D").ColumnWidth = base_width * 2
    Columns("F").ColumnWidth = base_width * 2.5
    Columns("G").ColumnWidth = base_width * 2.5
    Columns("H").ColumnWidth = base_width * 2.5
    
    ' add the values to the empty columns
    Range("A1").Value = UCase("ordem")
    Range("B1").Value = UCase("prioridade")
    Range("C1").Value = UCase("linha")
    Range("D1").Value = UCase("operação")
    Range("E1").Value = UCase("ativo")
    Range("F1").Value = UCase("tipo de manutenção")
    Range("G1").Value = UCase("natureza do serviço")
    Range("H1").Value = UCase("tempo estimado")
    
    ' center the table titles
    Range(title_range).VerticalAlignment = xlCenter
    Range(title_range).HorizontalAlignment = xlCenter
    
    ' center table content
    Range(non_title_range).VerticalAlignment = xlCenter
    Range(non_title_range).HorizontalAlignment = xlCenter
    
    
    ' change the font weight of the table titles
    Range(title_range).Font.Bold = True
    
    ' add conditional formatting rules
    ' makes empty cells blank
    With Worksheets(1).Range(empty_cells_range).FormatConditions _
        .Add(xlBlanksCondition)
        With .Borders
            .Color = RGB(255, 255, 255)
        End With
    End With
    
    ' makes the title cells black with a white bold text
    With Worksheets(1).Range(title_range).FormatConditions _
        .Add(xlNoBlanksCondition)
        With .Interior
            .ColorIndex = 1
        End With
        
        With .Font
            .Bold = True
            .ColorIndex = 2
        End With
    End With
    
    ' adds a black borders to filled cells that are not the title cells
    With Worksheets(1).Range(non_title_range).FormatConditions _
        .Add(xlNoBlanksCondition)
        With .Borders
            .Color = RGB(0, 0, 0)
        End With
    End With
End Sub

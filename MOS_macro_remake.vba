

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


Sub CreateDummyMOs()
    ' Create some dummy data of Maintence Orders (M.Os) to test code

    ' Declare variables
    Dim dummy_mos_to_generate As Integer
    Dim nature_pos As Integer
    Dim line_pos As Integer
    Dim op_pos As Integer
    Dim type_pos As Integer
    Dim active_pos As Integer
    Dim priority_pos As Integer
    Dim etd_pos As Integer
    
    Dim nature_type_no As Integer
    Dim line_type_no As Integer
    Dim op_type_no As Integer
    Dim type_type_no As Integer
    Dim active_type_no As Integer
    Dim etd_type As Integer
    
    ' Set the number of MOs to generate
    dummy_mos_to_generate = 20
    
    ' Set the colum to write each type of dummy data
    priority_pos = 2
    line_pos = 3
    op_pos = 4
    active_pos = 5
    type_pos = 6
    nature_pos = 7
    etd_pos = 8

    For i = 2 To dummy_mos_to_generate
    
        ' Create random MOs nubmers that start with 22
        Cells(i, 1) = "22" & Int((9999 - 1000 + 1) * Math.Rnd() + 1000)
    
        ' Get a random number that will be the MO nature type
        nature_type_no = Int((2 - 0) * Math.Rnd() + 1)
    
        ' set random MOs types
        If nature_type_no = 1 Then
            Cells(i, nature_pos) = "ELE"
            
        Else
            Cells(i, nature_pos) = "MEC"
        End If
        
        ' Get a random number that will be the MO line type
        line_type_no = Int((8 - 0) * Math.Rnd() + 1)
        
        If line_type_no = 1 Then
            Cells(i, line_pos) = "T XBB"
        
        ElseIf line_type_no = 2 Then
            Cells(i, line_pos) = "T HHA"
        
        ElseIf line_type_no = 3 Then
            Cells(i, line_pos) = "T X52"
        
        ElseIf line_type_no = 4 Then
            Cells(i, line_pos) = "PEM 001"
        
        ElseIf line_type_no = 5 Then
            Cells(i, line_pos) = "PEM 002"
            
        ElseIf line_type_no = 6 Then
            Cells(i, line_pos) = "PEM 003"
        
        ElseIf line_type_no = 7 Then
            Cells(i, line_pos) = "PEM 004"
        
        ElseIf line_type_no = 8 Then
            Cells(i, line_pos) = "PET 001"
        
        Else
            Cells(i, line_pos) = "PET 002"
        
        End If
        
        ' Get a random number that will be the MO op type
        op_type_no = Int((5 - 0) * Math.Rnd() + 1)
        
        If line_type_no = 1 Then
            Cells(i, op_pos) = "OP 5"
        
        ElseIf line_type_no = 2 Then
            Cells(i, op_pos) = "OP 10"
        
        ElseIf line_type_no = 3 Then
            Cells(i, op_pos) = "OP 15"
        
        ElseIf line_type_no = 4 Then
            Cells(i, op_pos) = "OP A/B"
        
        ElseIf line_type_no = 5 Then
            Cells(i, op_pos) = "OP 100/110"
        
        Else
            Cells(i, op_pos) = "CARRO TRANS. FER."
        
        End If
        
        ' Get the op type that defines the active type
        active_type = Cells(i, op_pos).Value
        active_type_no = Int((2 - 0) * Math.Rnd() + 1)
        
        If active_type = "CARRO TRANS. FER." Then
            Cells(i, active_pos) = "CTF"
               
        ElseIf active_type_no = 1 Then
            Cells(i, active_pos) = "ROB"
        
        ElseIf active_type_no = 2 Then
            Cells(i, active_pos) = "DSP"
        
        Else
            Cells(i, active_pos) = "PRP"
            
        
        End If
        
        ' Get a random number that will be the MO type type
        type_type_no = Int((2 - 0) * Math.Rnd() + 1)
    
        ' Create random MOs types
        If type_type_no = 1 Then
            Cells(i, type_pos) = "PREVENTIVA"
            Cells(i, priority_pos) = "A"
        Else
            Cells(i, type_pos) = "CORRETIVA P."
            Cells(i, priority_pos) = "B"
        End If
        
        ' Get a random number that will be the MO etd type
        etd_type_no = Int((3 - 0) * Math.Rnd() + 1)
    
        ' Create random MOs etds
        If etd_type_no = 1 Then
            Cells(i, etd_pos) = "0.85"
            
        ElseIf etd_type_no = 2 Then
            Cells(i, etd_pos) = "1.00"

        Else
            Cells(i, etd_pos) = "0.50"
            
        End If
    
    Next i


End Sub

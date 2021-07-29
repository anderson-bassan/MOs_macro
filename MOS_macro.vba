' TO-DO
' * improve find algorithm
' * make add mo dummy proof
' * add reprogrammed times to add mo
' * add remove repeated mos to clean up
' * add function to create back-ups csv files
'


Sub FormatSpreadsheet()
    ' Declare the ranges
    Dim title_range As String
    Dim non_title_range As String
    Dim empty_cells_range As String
    
    ' Set the base width of the columns
    Dim base_width As Double
    
    base_width = 8.43
    
    ' Set the ranges that will be used
    title_range = "A1:H1"
    non_title_range = "A2:H999"
    empty_cells_range = "A:ZZ"
    
    ' Change the width of the used columns
    Columns("B").ColumnWidth = base_width * 2
    Columns("D").ColumnWidth = base_width * 2
    Columns("F").ColumnWidth = base_width * 2.5
    Columns("G").ColumnWidth = base_width * 2.5
    Columns("H").ColumnWidth = base_width * 2.5
    
    ' Add the values to the empty columns
    Range("A1").Value = UCase("ordem")
    Range("B1").Value = UCase("prioridade")
    Range("C1").Value = UCase("linha")
    Range("D1").Value = UCase("operação")
    Range("E1").Value = UCase("ativo")
    Range("F1").Value = UCase("tipo de manutenção")
    Range("G1").Value = UCase("natureza do serviço")
    Range("H1").Value = UCase("tempo estimado")
    
    ' Center the table titles
    Range(title_range).VerticalAlignment = xlCenter
    Range(title_range).HorizontalAlignment = xlCenter
    
    ' Center table content
    Range(non_title_range).VerticalAlignment = xlCenter
    Range(non_title_range).HorizontalAlignment = xlCenter
    
    
    ' Change the font weight of the table titles
    Range(title_range).Font.Bold = True
    
    ' Add conditional formatting rules
    ' Makes empty cells blank
    With Worksheets(1).Range(empty_cells_range).FormatConditions _
        .Add(xlBlanksCondition)
        With .Borders
            .Color = RGB(255, 255, 255)
        End With
    End With
    
    ' Makes the title cells black with a white bold text
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
    
    ' Adds black borders to filled cells that are not the title cells
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
    
        ' Set random MOs types
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


Function FindLastMOIndex() As Integer
    ' Find the last MO of the spreadsheet and returns it
        
    ' Declare variables
    Dim last_mo_index As Integer
    Dim current_mo As Long
    Dim total_cells As Long
    
    ' Set the number of cells to check
    total_cells = 2500
    
    ' Run throught the first X (total_cells) cells until it finds the one with content
    For i = total_cells To 2 Step -1
        If Not IsEmpty(Cells(i, 1)) And last_mo_index = 0 Then
            last_mo_index = i

        End If
    Next i
    
    If last_mo_index = 0 Then
        last_mo_index = 1
    End If
    
    ' Return the last MO index
    FindLastMOIndex = last_mo_index
        
End Function


Function FindLastMO() As String
    ' Find the last MO of the spreadsheet and returns it's OM no. cell
        
    FindLastMO = "A" & FindLastMOIndex
        
End Function


Sub DeleteEmptyMOs()
    ' Find and delete all empty cells
    
    ' Variable Declaration
    Dim last_cell As String
    
    ' Set the last cell, that will be used to search between all cells in the table by empty cells
    last_cell = "G" & FindLastMOIndex
    
    On Error Resume Next
    ' Deletes all lines with empty cells, except for ETD as it's optional
    Range("A1", last_cell).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    On Error GoTo 0

End Sub


Sub SortMOs()
    ' Sorts Maintence Orders

    ' Sorts all MOs by MO no.
    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        ' .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SetRange Range("A1:H25000")
        .Header = xlYes
        .Apply
    End With
    
    ' Avoid stacking up sort key informations
    ActiveSheet.Sort.SortFields.Clear

End Sub


Sub CleanUpTable()
    ' Sorts the entire table and clean up empty cells

    DeleteEmptyMOs
    SortMOs
End Sub


Sub FindMOByNumber(mo_number)
    ' Declare variables
    Dim current_mo_value As Long
    Dim found_mo As Boolean
    Dim value_location As String
    
    ' Sets found mo to false, so it knows latter when it finds it
    found_mo = False
    
    ' Delete empty MOs and sorts the whole table
    CleanUpTable
    
    ' Loops through every MO and compare values, if it finds it selects the cell of the MO
    For i = 2 To FindLastMOIndex
        current_mo_value = Range("A" & i).Value
        
        If current_mo_value = mo_number Then
            Range("A" & i).EntireRow.Select
            found_mo = True
            
        End If
    Next i
    
    ' If no MO is found then pop-up a message box saying such
    If found_mo = False Then
        MsgBox ("OM não encontrada...")
        
    End If
End Sub


Sub FindMO()
    ' MO by the value given by the user

    ' Declare variables
    Dim mo_number As String

    ' Gets the number from the user
    mo_number = InputBox("Qual o número da OM? ", "Número da OM")

    If mo_number <> "" Then
        ' Search MO by the value
        FindMOByNumber (mo_number)
    
    Else
        ' Informs the user that it was not possible to find the OM
        If found_mo = False Then
            MsgBox ("Não foi possível buscar a OM...")
        
        End If
    End If
End Sub


Function DoesMOExist(mo_number As String)
    ' Declare variables
    Dim last_mo As Integer
    Dim mo_exist As Boolean
    
    ' Set the mo_exist to false and the last row in the table
    mo_exist = False
    last_mo_index = FindLastMOIndex
    
    ' Go through all cells verifying if the mo exist
    For i = last_mo_index To 1 Step -1
        ' If it find the mo number it does exist, else it still dosen't exist
        If mo_number = Cells(i, 1).Value Then
            mo_exist = True
            
        End If
    Next i
    
    ' Returns if the mo exist
    DoesMOExist = mo_exist
    
End Function


Sub AddMO()
    ' Add a MO to the list, then delete empty cells and sorts it
    
    ' Declare variables
    Dim new_mo_row As Long
    
    Dim mo_number As String
    Dim mo_priority As String
    Dim mo_line As String
    Dim mo_op As String
    Dim mo_active As String
    Dim mo_type As String
    Dim mo_nature As String
    Dim mo_etd As String
    
    ' Get MO number
    mo_number = InputBox("número da O.M.: ", "Número da OM")
    mo_priority = InputBox("prioridade da O.M.: ", "prioridade da OM")
    mo_line = InputBox("linha da O.M.: ", "linha da OM")
    mo_op = InputBox("operação da O.M.: ", "operação da OM")
    mo_active = InputBox("ativo da O.M.: ", "ativo da OM")
    mo_type = InputBox("tipo de manutenção da O.M.: ", "tipo de manutenção da OM")
    mo_nature = InputBox("natureza de serviço: ", "natureza de serviço")
    mo_etd = InputBox("tempo estimado(opcional): ", "tempo estimado")

    If Not DoesMOExist(mo_number) Then
        ' Add Mo
        
        ' Select the cell where the MO will be add
        If CLng(mo_number) <> 0 And mo_priority <> "" And mo_line <> "" And mo_op <> "" And mo_active <> "" And mo_type <> "" And mo_nature <> "" Then
            new_mo_row = FindLastMOIndex + 1
            Range("A" & new_mo_row).Value = mo_number
            Range("B" & new_mo_row).Value = mo_priority
            Range("C" & new_mo_row).Value = mo_line
            Range("D" & new_mo_row).Value = mo_op
            Range("E" & new_mo_row).Value = mo_active
            Range("F" & new_mo_row).Value = mo_type
            Range("G" & new_mo_row).Value = mo_nature
    
            If mo_etd = "" Then
                Range("H" & new_mo_row).Value = "N/A"
                
            Else
                Range("H" & new_mo_row).Value = mo_etd
                
            End If
            
            FindMOByNumber (mo_number)
        
        Else
            ' Shows a message in case there's empty required fields
            MsgBox ("Alguns campos obrigatórios estão vazios...")
        
        End If

    Else
        ' Insert Reprogrammed
        MsgBox ("A O.M. já existe... ")
    
    End If
       
    ' Delete empty MOs and sort them
    CleanUpTable
    
End Sub


Sub DelMO()
    ' Delete a MO from the list, then delete empty cells and sorts it
    
    ' Declare Variables
    Dim mo_number As String
    Dim current_mo_value As Long
    Dim found_mo As Boolean
    Dim value_location As String
        
    ' Set mo_number to the "text box" number
    mo_number = InputBox("Número da O.M.: ", "número da O.M.")
    
    ' Set found_mo to false to check if it was found later
    found_mo = False
    
    ' Loops through every MO and compare values, if it finds it deletes the cell of the MO
    For i = 2 To FindLastMOIndex
        current_mo_value = Range("A" & i).Value
        
        If current_mo_value = CLng(mo_number) Then
            ' Ask the user if he/she/it is sure that he/she/it wants to delte the MO
            del_mo_answer = MsgBox("Você tem certeza? ", vbQuestion + vbYesNo + vbDefaultButton2, "Você tem certeza? ")
            
            ' Deletes when users confirm
            If del_mo_answer = vbYes Then
                Range("A" & i).EntireRow.Delete
                found_mo = True
                
            End If
            
        End If
    Next i
    
    ' Checks if the MO was found, otherwise sends a message
    If Not found_mo Or mo_number = "" Then
        MsgBox ("Não foi possível deletar a O.M. ...")
    End If
    
    ' Delete empty MOs and sort them
    CleanUpTable
    
End Sub


Sub TestSub()
    ' Tries to find where the bug (id: 7) is occurring
       
    ' FormatSpreadsheet: OK
    ' CreateDummyMOs: OK
    ' FindLastMOIndex: OK
    ' FindLastMO: OK
    ' DeleteEmptyMOs: OK
    ' SortMOs: NOT OK
    ' FindMO: OK obs.: it didn't worked before deactivating sortMOs
    ' DoesMOExist: OK
    ' AddMO: OK
    ' DelMO: OK
   
    ' Trying to fix the bug in sort function
    ' Works without the repeated line: OK
    ' works fine, but the bugs still happens after saving
    ' Works with the clear sort references afterwards: OK
    
    ' Using the changed Sub in the previous disabled code
    ' Works with the changes: OK
End Sub



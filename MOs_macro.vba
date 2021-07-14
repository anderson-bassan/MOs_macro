' TO-DO
' * add function to create back-ups csv files
' * add text-boxes instead of excel cells?
' * improve table to make double click sorts possible
' * when sorting MOs by double click show a graphic legend
' * implement quick find to FindLastOM and FindLastOMIndex
'
' TO-FIX
' * update sort MOs to sort the entire MO line and not just the MO no.


Sub CreateDummyMOs()
    ' Create some dummy data of Maintence Orders (M.Os) to test code

    ' Declare variables
    Dim dummy_mo_to_generate As Integer
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
    dummy_mo_to_generate = 20
    
    ' Set the colum to write each type of dummy data
    priority_pos = 2
    line_pos = 3
    op_pos = 4
    active_pos = 5
    type_pos = 6
    nature_pos = 7
    etd_pos = 8

    For i = 2 To dummy_mo_to_generate
    
        ' Create random MOs nubmers that start with 22
        Cells(i, 1) = "22" & Int((9999 - 1000 + 1) * Math.Rnd() + 1000)
    
        ' Get a random number that will be the MO nature type
        nature_type_no = Int((2 - 0) * Math.Rnd() + 1)
    
        ' Create random MOs types
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
            Cells(i, op_pos) = "op 5"
        
        ElseIf line_type_no = 2 Then
            Cells(i, op_pos) = "op 10"
        
        ElseIf line_type_no = 3 Then
            Cells(i, op_pos) = "op 15"
        
        ElseIf line_type_no = 4 Then
            Cells(i, op_pos) = "op A/B"
        
        ElseIf line_type_no = 5 Then
            Cells(i, op_pos) = "op 100/110"
        
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
            Cells(i, type_pos) = "PREVENTIVE"
            Cells(i, priority_pos) = "A"
        Else
            Cells(i, type_pos) = "P. CORRETIVE"
            Cells(i, priority_pos) = "B"
        End If
        
        ' Get a random number that will be the MO etd type
        etd_type_no = Int((3 - 0) * Math.Rnd() + 1)
    
        ' Create random MOs types
        If etd_type_no = 1 Then
            Cells(i, etd_pos) = "0.85"
            
        ElseIf etd_type_no = 2 Then
            Cells(i, etd_pos) = "1.00"

        Else
            Cells(i, etd_pos) = "0.50"
        End If
    
    Next i


End Sub


Function FindLastMO() As String
    ' Find the last MO of the spreadsheet and returns it
        
    ' Declare variables
    Dim last_mo As String
    Dim current_mo As Long
    Dim total_cells As Long
    
    ' Set the number of cells to check
    total_cells = 2500
    
    ' Run throught the first X (total_cells) cells until it finds the one with content
    For i = total_cells To 2 Step -1
        If Not IsEmpty(Cells(i, 1)) And last_mo = "" Then
            last_mo = "A" & i
        End If
    Next i
    
    ' Return the last MO index
    FindLastMO = last_mo
        
End Function


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
    
    ' Return the last MO index
    FindLastMOIndex = last_mo_index
        
End Function


Sub DeleteEmptyMOs(last_cell)
    ' Find and delete all empty cells
    
    On Error Resume Next
    ' Selects all MOs blanks and deletes them
    Range("A1", last_cell).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    Debug.Print "inside delteemptymos; last_cell: " & last_cell
    
    On Error GoTo 0

End Sub


Sub SortMOs(last_mo)
    'Sort Maintence Orders

    ' Sort all MOs
    Range("A1", last_mo).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

End Sub


Sub FindMO()
    ' Delete empty cells and sorts all them and search a MO by number

    ' Declare variables
    Dim mo_number As Long
    Dim current_mo_value As Long
    Dim found_mo As Boolean
    Dim value_location As String

    ' Set the location to retrive values from and to clear later
    value_location = "J2"

    ' Get the MO no. and set found_mo to false so it knows to pop-up the msgbox later in
    ' case it is not found
    mo_number = Range(value_location).Value
    found_mo = False

    ' Delete empty MOs and sort them
    DeleteEmptyMOs (FindLastMO)
    ' SortMOs (FindLastMO)
    
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
        MsgBox ("MO not found...")
    End If
    
    ' Clear "search box" after searching
    Range(value_location).Value = ""
    
End Sub


Sub AddMO()
    ' Add a MO to the list, then delete empty cells and sorts it
    
    ' Declare variables
    Dim new_mo_row As Long
    
    Dim mo_number As Long
    Dim mo_priority As String
    Dim mo_line As String
    Dim mo_op As String
    Dim mo_active As String
    Dim mo_type As String
    Dim mo_nature As String
    Dim mo_etd As String
    
    Dim value_location As String
    Dim priority_location As String
    Dim line_location As String
    Dim op_location As String
    Dim active_location As String
    Dim type_location As String
    Dim nature_location As String
    Dim etd_location As String
    
    ' Set the location to retrive values from and to clear later
    value_location = "J4"
    priority_location = "K4"
    line_location = "L4"
    op_location = "M4"
    active_location = "N4"
    type_location = "O4"
    nature_location = "P4"
    etd_location = "Q4"
    
    ' Get MO number
    mo_number = Range(value_location).Value
    mo_priority = Range(priority_location).Value
    mo_line = Range(line_location).Value
    mo_op = Range(op_location).Value
    mo_active = Range(active_location).Value
    mo_type = Range(type_location).Value
    mo_nature = Range(nature_location).Value
    mo_etd = Range(etd_location).Value
    
    ' Select the cell where the MO will be add
    If mo_number <> 0 And mo_priority <> "" And mo_line <> "" And mo_op <> "" And mo_active <> "" And mo_type <> "" And mo_nature <> "" Then
        new_mo_row = FindLastMOIndex + 1
        Range("A" & new_mo_row).Value = mo_number
        Range("B" & new_mo_row).Value = mo_priority
        Range("C" & new_mo_row).Value = mo_line
        Range("D" & new_mo_row).Value = mo_op
        Range("E" & new_mo_row).Value = mo_active
        Range("F" & new_mo_row).Value = mo_type
        Range("G" & new_mo_row).Value = mo_nature

    
    Else
        MsgBox ("Some required fields are empty...")
    
    End If
    
    If mo_etd = "" Then
        Range("H" & new_mo_row).Value = "N/A"
    Else
        Range("H" & new_mo_row).Value = mo_etd
    End If
    
    ' Clean up "text box"
    Range(value_location, etd_location).Value = ""
    
    ' Delete empty MOs and sort them
    DeleteEmptyMOs (FindLastMO)
    ' SortMOs (FindLastMO)
    
End Sub


Sub DelMO()
    ' Delete a MO from the list, then delete empty cells and sorts it
    
    ' Declare Variables
    Dim mo_number As Long
    Dim current_mo_value As Long
    Dim found_mo As Boolean
    Dim value_location As String
    
    ' Set the location to retrive values from and to clear later
    value_location = "J6"
    
    ' Set mo_number to the "text box" number
    mo_number = Range(value_location).Value
    
    ' Set found_mo to false to check if it was found later
    found_mo = False
    
    ' Loops through every MO and compare values, if it finds it deletes the cell of the MO
    For i = 2 To FindLastMOIndex
        current_mo_value = Range("A" & i).Value
        Debug.Print (current_mo_value)
        
        If current_mo_value = mo_number Then
            ' Ask the user if he/she/it is sure that he/she/it wants to delte the MO
            del_mo_answer = MsgBox("Are you sure? ", vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?")
            
            ' Deletes when users confirm
            If del_mo_answer = vbYes Then
                Range("A" & i).EntireRow.Delete
                found_mo = True
                
            End If
            
        End If
    Next i
    
    ' Checks if the MO was found, otherwise sends a message
    If Not found_mo Then
        MsgBox ("MO was not found...")
    End If
    
    ' Clears the "text box"
    Range(value_location).Value = ""
    
    ' Delete empty MOs and sort them
    DeleteEmptyMOs (FindLastMO)
    ' SortMOs (FindLastMO)
    
End Sub


Sub TestMacro()
    Range("J4", "L4").Select
    
End Sub

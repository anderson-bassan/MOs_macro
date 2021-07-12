' TO-DO
' * add the AddMO Sub
' * add the RemoveMo Sub
'
'
' TO-FIX
' * when running the DeleteEmptyMOs with no empty MOs returns an error.
'


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
        If Not IsEmpty(Cells(i, 1)) And last_mo = "" Then
            last_mo_index = i
        End If
    Next i
    
    ' Return the last MO index
    FindLastMOIndex = last_mo_index
        
End Function


Sub CreateRandomMOs()
    ' Create some test data of Maintence Orders (M.Os) to test code

    ' Create random MOs nubmers that start with 22
    For i = 2 To 20
        Cells(i, 1) = "22" & Int((9999 - 1000 + 1) * Math.Rnd() + 1000)
    Next i

End Sub


Sub DeleteEmptyMOs(last_cell)
    ' Find and delete all empty cells
    
    ' Selects all MOs blanks and deletes them
    Range("A1", last_cell).SpecialCells(xlCellTypeBlanks).Delete

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

    ' Get the MO no. and set found_mo to false so it knows to pop-up the msgbox later in
    ' case it is not found
    mo_number = Range("C2").Value
    found_mo = False

    ' Delete empty MOs and sort them
    DeleteEmptyMOs (FindLastMO)
    SortMOs (FindLastMO)
    
    ' Loops through every MO and compare values, if it finds it selects the cell of the MO
    For i = 2 To FindLastMOIndex
        current_mo_value = Range("A" & i).Value
        If current_mo_value = mo_number Then
            Range("A" & i).Select
            found_mo = True
        End If
    Next i
    
    ' If no MO is found then pop-up a message box saying such
    If found_mo = False Then
        MsgBox ("MO not found...")
    End If
    
    ' Clear "search box" after searching
    Range("C2").Value = ""
    
End Sub

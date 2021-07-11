Function FindLastMO() As Integer
    ' Find the last MO of the spreadsheet and returns it
        
    Dim last_mo As Integer
    Dim current_mo As Long
    
    For i = 25000 To 2 Step -1
        If Not IsEmpty(Cells(i, 1)) And last_mo = 0 Then
            last_mo = i
        End If
        
    Next i
    
    FindLastMO = last_mo
        
End Function
Sub CreateRandomMOs()
    ' Create some test data of Maintence Orders (M.Os) to test code

    For i = 2 To 20
        Cells(i, 1) = "22" & Int((9999 - 1000 + 1) * Math.Rnd() + 1000)
    
    Next i
End Sub
Sub DeleteEmptyCells(last_cell)
    ' Find and delete all empty cells
    
    Range("A1", "A" & last_cell).SpecialCells(xlCellTypeBlanks).Delete
    ' To fix: when run with no empty cells generate an error.
End Sub
Sub SortMOs(last_mo)
    'Sort Maintence Orders

    Range("A1", "A" & last_mo).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub
Sub FindMO()
    ' Delete empty cells and sorts all them and (TODO) search a MO by number

    Dim mo_number As Long
    Dim current_mo_value As Long
    Dim found_mo As Boolean

    mo_number = Range("C2").Value
    found_mo = False

    DeleteEmptyCells (FindLastMO)
    SortMOs (FindLastMO)
    
    For i = 2 To FindLastMO
        current_mo_value = Range("A" & i).Value
        If current_mo_value = mo_number Then
            Range("A" & i).Select
            found_mo = True
        End If
        
    Next i
    
    If found_mo = False Then
        MsgBox ("MO not found...")
    End If
    
    
End Sub

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

    For i = 2 To 1090
        Cells(i, 1) = "22" & Int((9999 - 1000 + 1) * Math.Rnd() + 1000)
    
    Next i
End Sub
Sub DummySub()
    Debug.Print FindLastMO
    
End Sub

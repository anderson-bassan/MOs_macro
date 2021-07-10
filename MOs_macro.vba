Sub CreateRandomMOs()
    ' Create some test data of Maintence Orders (M.Os) to test code

    For i = 2 To 1090
        Cells(i, 1) = "22" & Int((9999 - 1000 + 1) * Math.Rnd() + 1000)
    
    Next i
End Sub

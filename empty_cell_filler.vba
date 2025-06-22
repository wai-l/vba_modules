Sub emptycellfiller()

    ' the column after the last empty column
    Range("N2").Select

    Do Until IsEmpty(ActiveCell.Value)

        If IsEmpty(ActiveCell.Offset(0, -1).Value) Then 
            ActiveCell.Offset(0, -1).Values = ActiveCell.Offset(-1, -1).Value
            ActiveCell.Offset(0, -2).Values = ActiveCell.Offset(-1, -2).Value
            ......
        End If
        ActiveCell.Offset(1, 0).Select
    Loop

End Sub
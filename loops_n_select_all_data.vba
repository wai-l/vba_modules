' Do Until + If
Do Until IsEmpty(ActiveCell.Value)

    If IsEmpty(ActiveCell.Offset(0, -1).Value) Then 
        ActiveCell.Offset(0, -1).Values = ActiveCell.Offset(-1, -1).Value
        ActiveCell.Offset(0, -2).Values = ActiveCell.Offset(-1, -2).Value
        ......
    End If
    ActiveCell.Offset(1, 0).Select
Loop

' Another Do Unil Loop
Range ("A2").Select

Do Until IsEmpty(ActiveCell.Value)
    If ActiveCell.Value = 0 Then 
        ActiveCell.EntireRow.Delete
    Else: ActiveCell.Offset(1, 0).Select
    End If
Loop

' One more Do Until Loop

Range("A6").Select

Do Until IsEmpty(ActiveCell.Value)
    SourcePath = SourceFolder & ActiveCell.Value
    PastePath = EndFolder & ActiveCell.Value

    FileCopy SourcePath, PastePath

    ActiveCell.Offset(1, 0).Select
Loop


' Select all rows by concating column header & rows.count
Range("A9:A" & Rows.Count).ClearContent

' Used Range
ActiveSheet.UsedRange.Select

' For Loop
Set Cell = Range("A9")

For Each i in TextArray
    Cell.Value = i
    Set Cell = Cell.Offset(1, 0)
Next i

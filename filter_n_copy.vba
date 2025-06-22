Sub Copy_n_paste_example()

' filtering
    ActiveSheet.UsedRange.AutoFilter Field:= 10, Criteria1:="Filtering Value"

' filtering with wildcard
    ActiveSheet.UsedRange.AutoFilter Field:= 11, Criteria1:= Array( _ 
    "*" & "some texts" & "*" _ 
    ), Operator:=xlFilterValue

' Format Date
    Columns("A:A").Select
    Selection.NumberFormat = "yyyy-mm-dd;@"

' Select filtered and copy
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets.Add After:= ActiveSheet
    ActiveSheet.Paste

End Sub

' Another way to filter wildcard
Dim WildCard As String
WildCard = "row: " & "*" & "This is some text" & "######" & " amount " & "*"

Range("A2").Select

Do Until IsEmpty(ActiveCell.Value)
    If ActiveCell.Value Like WildCard Then 
        ActiveCell.Offset(0, 1).Value = "TRUE"
        Else
        ActiveCell.Offset(0, 1).Value = "FALSE"
    End If
    ActiveCell.Offset(1, 0).Select
Loop
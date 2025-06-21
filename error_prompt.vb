On Error Resume Next
    DueDate = Application.WorksheetFunction.VLookup(AccNo, Worksheets("Tab1").Range("B:G"), 6, False)

If Err.Number <> 0 Then
    DueDate = ""
    Err.Clear

End If

On Error GoTo 0

If DueDate = "00:00:00", Then
    DueDate = Date
Else
    DueDate = DueDate

End If
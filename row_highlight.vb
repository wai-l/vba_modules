Sub highlighter()
'
' highlighter Macro
'

    set ws = ActiveSheet
' get table range
    Dim selected_range AS Range
    If TypeName(Selection) = "Range" Then
        Set selected_range = Selection
    Else
        MsgBox "Please select a range first." 
        Exit Sub
    End If

    
    id_col = InputBox("Enter the column letter for the ID (e.g. A): ", "ID Column", "A")
    IF id_col = "" Then
        MsgBox "No column selected. Exiting."
        Exit Sub
    End If
    
    id_col_num = Range(id_col & "1").Column

    first_row = selected_range.Row + 1
    last_row = first_row + selected_range.Rows.Count - 2
    current_row = first_row

    Range(selected_range.address(True, True)).Select

    ' identify groups of IDs that needs to be highlighted, we are only highlighting every 2nd group
    group_count = 0
    Do While current_row <= last_row
        group_start = current_row
        current_id = Cells(current_row, id_col_num).Value
        ' Find the end of the current ID group
        Do While current_row <= last_row And Cells(current_row, id_col_num).Value = current_id
            current_row = current_row + 1
        Loop

        ' Highlight the group if it's an odd group
        group_count = group_count + 1
        odd_group = (group_count Mod 2 = 1)
        even_group = Not odd_group

        Dim start_cell As Range
        Dim end_cell As Range
        Dim row_range As Range
        If odd_group Then
            For r = group_start To current_row - 1
                Set start_cell = Cells(r, 1)
                Set end_cell = Cells(r, selected_range.Columns.Count)
                Set row_range = ws.Range(start_cell, end_cell)
                With row_range.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = 0
                End With
            Next r
        ElseIf even_group Then
            For r = group_start To current_row - 1
                Set start_cell = Cells(r, 1)
                Set end_cell = Cells(r, selected_range.Columns.Count)
                Set row_range = ws.Range(start_cell, end_cell)
                With row_range.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 00
                End With
            Next r
        End If
    Loop

End Sub

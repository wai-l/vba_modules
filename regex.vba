Range("A2").Select

Dim OrgText As String
Dim TextBefore As String
Dim AccountNo_Posit As Variant

TextBefore = "The account number: "

Range("A2").Select

Do Until IsEmpty(ActiveCell.Value)
    OrgText = ActiveCell.Value
    AccountNo_Posit = InStr(1, OrgText, TextBefore, vbTextCompare)
    If ActiveCell.Offset(0, 1).Value = True Then
        Active.Cell.Offset(0, 2).Value = Mid(OrgText, AccountNo_Posit + Len(TextBefore), 7)
    Else 
        ActiveCell.Offset(0, 2).Value = ""
    End If

ActiveCell.Offset(1, 0).Select

Loop
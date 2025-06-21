Sub LineSplitter()

    Dim TextArray As Variant
    Dim Keyword As String
    Dim Cell As Range

    Range("A9:A" & Rows.Count).ClearContents

    Keyword = Range("B2").Value
    KeywordConcat = Keyword & "-" & Keyword
    KeywordSplit = Keyword & "-"
    Text = Replace(Range("B1"), Keyword, KeywordConcat)
    TextArray = Split(Text, KeywordSplit)

    Set Cell = Range("A9")

    For Each i in TextArray
        Cell.Value = i
        Set Cell = Cell.Offset(1, 0)
    Next i

End Sub
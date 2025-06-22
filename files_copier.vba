Sub File_Copier()

Dim SourceFolder, EndFolder, FileName As String

SourceFolder = Range("B1").Value
EndFolder = Range("B2").Value

Range("A6").Select

Do Until IsEmpty(ActiveCell.Value)
    SourcePath = SourceFolder & ActiveCell.Value
    PastePath = EndFolder & ActiveCell.Value

    FileCopy SourcePath, PastePath

    ActiveCell.Offset(1, 0).Select
Loop


End Sub
Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  MoreFasterCode (True)
  With ThisWorkbook.Worksheets("Feuil1")
    Dim Cellule As Range: Dim Ligne As Long
    Ligne = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    For Each Cellule In Target.Cells
        .Cells(Ligne, 1).Value = Format(Now(), "YYYY-MM-DD hh:mm:ss")
        .Cells(Ligne, 2).Value = Environ("USERNAME")
        .Cells(Ligne, 3).Value = Cellule.Row
        .Cells(Ligne, 4).Value = Cellule.Column
        .Cells(Ligne, 5).Value = Cellule.Value
        Ligne = Ligne + 1
    Next Cellule: Set Cellule = Nothing
  End With
  MoreFasterCode (False)
End Sub

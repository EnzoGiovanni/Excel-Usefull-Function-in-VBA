Function FindAllElt(ByRef Element As Variant, ByRef zone As Range) As Range

    'Preparation of sheet
    On Error Resume Next
        zone.Worksheet.ShowAllData
    On Error GoTo 0
    With zone.Worksheet
        .Rows.EntireRow.Hidden = False
        .Columns.EntireColumn.Hidden = False
    End With
    
    'finding all elements
    Dim Elt, Aire As Range: Dim FirstAdresse As String
    For Each Aire In zone.Areas
        With Aire
            Set Elt = .Find(Element, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=True)
            If Not Elt Is Nothing Then
                FirstAdresse = Elt.Address()
                Do
                    If FindAllElt Is Nothing Then
                        Set FindAllElt = Elt
                    Else
                        Set FindAllElt = .Application.Union(FindAllElt, Elt)
                    End If
                    Set Elt = .FindNext(Elt)
                Loop While Elt.Address <> FirstAdresse And Not Elt Is Nothing
            End If
        End With
    Next Aire: Set Aire = Nothing: Set Elt = Nothing
    
End Function

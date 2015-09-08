Function MoreFasterCode(ByRef Top As Boolean)
    With Application
        If Top Then
            .DisplayAlerts = False: .ScreenUpdating = False: .EnableEvents = False: .Calculation = xlManual
        Else
            .Calculation = xlAutomatic: .EnableEvents = True: .ScreenUpdating = True: .DisplayAlerts = True
        End If
    End With
End Function

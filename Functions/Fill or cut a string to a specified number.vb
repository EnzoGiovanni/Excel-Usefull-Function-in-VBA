'Fill or cut a string to a specified number
Function CompletingString(ByRef StrIn As String, ByRef NbCarMax As Long) As String
    Dim NbCar As Long: NbCar = Len(StrIn)
    CompletingString = StrIn
    If NbCar <= NbCarMax Then
        Dim i As Long
        For i = NbCar To NbCarMax - 1 Step 1
            CompletingString = CompletingString & " "
        Next i
    Else
        CompletingString = Left(CompletingString, NbCarMax)
    End If
End Function

Function TotalEffectiveRate(amount As Double, MonthPay As Double, NbEch As Long) As Double
    'anount : Total amount of credit deducted all expenses related to credit
    'MonthPay : monthly payment
    
    'hypothesis rate
    TotalEffectiveRate = 0.01
    
    'Résolution
    Dim DeltaHypTx As Double
    DeltaHypTx = Poly(amount, MonthPay, TotalEffectiveRate, NbEch) / PolyPrime(amount, MonthPay, TotalEffectiveRate, NbEch)
    Do While Abs(DeltaHypTx) >= 0.000001
        TotalEffectiveRate = TotalEffectiveRate - DeltaHypTx
        DeltaHypTx = Poly(amount, MonthPay, TotalEffectiveRate, NbEch) / PolyPrime(amount, MonthPay, TotalEffectiveRate, NbEch)
    Loop
End Function
Function Poly(ByRef amount As Double, ByRef MonthPay As Double, ByRef Tx As Double, ByRef NbEch As Long) As Double
    Poly = amount / MonthPay * Tx * ((1 + Tx) ^ NbEch) - ((1 + Tx) ^ NbEch) + 1
End Function
Function PolyPrime(ByRef amount As Double, ByRef MonthPay As Double, ByRef Tx As Double, ByRef NbEch As Long) As Double
    PolyPrime = amount / MonthPay * ((1 + Tx) ^ NbEch) + amount / MonthPay * NbEch * Tx * ((1 + Tx) ^ (NbEch - 1)) - NbEch * ((1 + Tx) ^ (NbEch - 1))
End Function

Function TotalEffectiveRate(amount As Double, MonthPay As Double, NbEch As Long) As Double
    'anount : Total amount of credit deducted all expenses related to credit
    'MonthPay : monthly payment
    'NbEch : Number of monthly payment
    
    'hypothesis rates
    TotalEffectiveRate = 0.01
    
    'Résolution
    Dim DeltaHypTx As Double
    DeltaHypTx = Poly(amount, MonthPay, TotalEffectiveRate, NbEch) / PolyPrime(amount, MonthPay, TotalEffectiveRate, NbEch)
    Do While Abs(DeltaHypTx) >= 0.0000001
        TotalEffectiveRate = TotalEffectiveRate - DeltaHypTx
        DeltaHypTx = (amount / MonthPay * TotalEffectiveRate * ((1 + TotalEffectiveRate) ^ NbEch) - ((1 + TotalEffectiveRate) ^ NbEch) + 1) _
                     / (amount / MonthPay * ((1 + TotalEffectiveRate) ^ NbEch) + amount / MonthPay * NbEch * TotalEffectiveRate * ((1 + TotalEffectiveRate) ^ (NbEch - 1)) - NbEch * ((1 + TotalEffectiveRate) ^ (NbEch - 1)))
    Loop
    
End Function

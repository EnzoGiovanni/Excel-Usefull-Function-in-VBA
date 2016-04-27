Function MonthlyRate(AnnualRate As Long)
    MonthlyRate = ((1 - AnnualRate) ^ (1 / 12)) - 1
End Function

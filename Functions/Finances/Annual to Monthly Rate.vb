Function MonthlyRate(AnnualRate As Currency) As Currency
    MonthlyRate = ((1 - AnnualRate) ^ (1 / 12)) - 1
End Function

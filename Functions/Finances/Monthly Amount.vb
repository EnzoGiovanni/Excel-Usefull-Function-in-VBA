Function MonthlyAmount(BorrowedCapital As Currency, NumberOfMaturity As Currency, MonthlyRate As Currency) As Currency
    MonthlyAmount = (BorrowedCapital * MonthlyRate) / (1 - ((1 + MonthlyRate) ^ (-NumberOfMaturity)))
End Function

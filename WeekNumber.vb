Function WeekNumber(d As Date)
    'Calcul du n° de semaine selon la norme ISO, norme europe
    Dim date_jeudi, date_4_janvier, date_lundi_semaine_1 As Date
    Dim Nb_jours, numero As Integer
    date_jeudi = DateAdd("d", 4 - Weekday(d, vbMonday), d)
    date_4_janvier = DateSerial(Year(date_jeudi), 1, 4)
    date_lundi_semaine_1 = DateAdd("d", 1 - Weekday(date_4_janvier, vbMonday), date_4_janvier)
    Nb_jours = Abs(DateDiff("d", date_lundi_semaine_1, date_jeudi, vbMonday))
    WeekNumber = Int(Nb_jours / 7) + 1
End Function

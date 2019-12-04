'traitement de reformatage des données du résultat d'une requette DB2 paginé
'###################################################################################################
Option Compare Database
'###################################################################################################
Sub LireFichierTexteParLigne()
'###################################################################################################
    Dim IndexFichier, i, k, NumLigneSeparation, NbPages As Integer: NumLigneSeparation = 0
    Dim Top, TopTitres, TopDebLi As Boolean: Top = False: TopTitres = False: TopDebLi = True
    Dim TitresDatas(), Lignes, ContenuLigne, FileIn, FileOut As String
    
    '===================================================================================================
    'Extraction SQL DB2 d'entrée
    FileIn = "C:\Users\368790\Downloads\carte agence 03102.txt"
    
    'Extraction SQL DB2 de sortie formaté CSV
    FileOut = "C:\Users\368790\Downloads\Out.txt"
    '===================================================================================================
    
    
       
    '===================================================================================================
    'Detecter les colonnes
    '===================================================================================================
    NumLigneSeparation = 0: k = 0
    '===================================================================================================
    IndexFichier = FreeFile()
    Open FileIn For Input As #IndexFichier 'ouvre le fichier
    Do While Not EOF(IndexFichier)
        '===================================================================================================
        Line Input #IndexFichier, ContenuLigne     ' lecture du fichier ligne par ligne: la variable "ContenuLigne" contient le contenu de la ligne active
        '===================================================================================================
        'Si ligne de séparation
        If (InStr(1, ContenuLigne, "PAGE", vbTextCompare) > 1) Then NumLigneSeparation = 0
        If (InStr(1, ContenuLigne, "  +--", vbTextCompare) > 1) Then NumLigneSeparation = NumLigneSeparation + 1
        '===================================================================================================
        'Si ligne de titres de colonne
        If ((InStr(1, ContenuLigne, "  ! ", vbTextCompare) > 1) And (NumLigneSeparation = 1)) Then
            '===================================================================================================
            Lignes = Split(ContenuLigne, "!")
            For i = LBound(Lignes) To UBound(Lignes)
                Lignes(i) = Trim(Lignes(i))
            Next i
            '===================================================================================================
            'Enregistrement des noms de colonnes
            For i = LBound(Lignes) To UBound(Lignes)
               If (Len(Lignes(i)) > 0) Then
                    If (Top) Then
                        If TopTitres Then If (TitresDatas(0) = Lignes(i)) Then Exit Do
                        ReDim Preserve TitresDatas(UBound(TitresDatas) + 1): If Not TopTitres Then TopTitres = True
                    Else
                        ReDim Preserve TitresDatas(0): Top = True
                    End If
                    TitresDatas(k) = Lignes(i): k = k + 1
               End If
            Next i
            '===================================================================================================
            NbPages = NbPages + 1
        End If
    '===================================================================================================
    Loop
    '===================================================================================================
    Close #IndexFichier ' ferme le fichier
    '===================================================================================================
    
    '===================================================================================================
    'Lire les enregistrements, les formater puis les ecrires dans un csv
    '===================================================================================================
    Dim Taille As Integer: Taille = UBound(TitresDatas) - LBound(TitresDatas) + 1
    Dim Rec(), RecOut As String
    Dim PageCourante As Integer: PageCourante = 0
    Dim DebLi, RecLig As Long: DebLi = 0: RecLig = 0
    NumLigneSeparation = 0
    '===================================================================================================
    Open FileIn For Input As #IndexFichier 'ouvre le fichier source
    IdxOutFile = FreeFile()
    Open FileOut For Output As #IdxOutFile  'ouvre le fichier destination
    '===================================================================================================
    'Ecrit le titre des colonnes dans le fichier de sorti
    RecOut = ""
    Dim iMin, iMax As Integer: iMin = LBound(TitresDatas()): iMax = UBound(TitresDatas())
    For i = iMin To iMax
        If i < iMax Then
            RecOut = RecOut + TitresDatas(i) + ";"
        Else
            RecOut = RecOut + TitresDatas(i)
        End If
    Next i
    Print #IdxOutFile, RecOut
    '===================================================================================================
    While Not EOF(IndexFichier) '
        '===================================================================================================
        Line Input #IndexFichier, ContenuLigne     ' lecture du fichier ligne par ligne
        '===================================================================================================
        'Si ligne de séparation
        If (InStr(1, ContenuLigne, "PAGE", vbTextCompare) > 1) Then NumLigneSeparation = 0
        If (InStr(1, ContenuLigne, "  +--", vbTextCompare) > 1) Then NumLigneSeparation = NumLigneSeparation + 1
        '===================================================================================================
        'Si ligne de titres de colonne
        If ((InStr(1, ContenuLigne, "  ! ", vbTextCompare) > 1) And (NumLigneSeparation = 1)) Then
            If PageCourange < NbPages Then
                PageCourante = PageCourante + 1
            Else
                PageCourante = 1
            End If
        End If
        '===================================================================================================
        'Si ligne de données
        If ((InStr(1, ContenuLigne, "_! ", vbTextCompare) > 1) And (NumLigneSeparation = 2)) Then
        
            'Nettoyage
            Lignes = Split(ContenuLigne, "!")
            For i = LBound(Lignes) To UBound(Lignes): Lignes(i) = Trim(Lignes(i)): Next i
            Lignes(0) = Split(Lignes(0), "_")(0)
                        
            'Calculer la variable DEBUT-LIGNE
            RecLig = CLng(Lignes(0))
            If TopDebLi Then DebLi = RecLig: TopDebLi = False
            
            'Transfert à Rec
            ReDim Preserve Rec(1 To Taille, DebLi To RecLig) '/!\ BUG /!\  ne pas mettre "1 To RecLig" mais "DEBUT-LIGNE to RecLig"
            For i = 1 To (UBound(Lignes) - 1)
                Rec(i, RecLig) = Lignes(i)
            Next i
            
            ' Maintenant , remplir la 2éme page
            
            
            
            
        End If
        '===================================================================================================
        'Si fin de page de renvoie alors décharger les données du tableau Rec et le vider
        '===================================================================================================
        If ((PageCourante = NbPages) And (InStr(1, ContenuLigne, "  +--", vbTextCompare) > 1) And (NumLigneSeparation = 3)) Then
            '===================================================================================================
            'Transférer REC dans le fichier de sortie
            Dim ColMin, ColMax As Integer
            For Lig = LBound(Rec(), 2) To UBound(Rec(), 2)
                RecOut = "": ColMin = LBound(Rec(), 1): ColMax = UBound(Rec(), 1)
                For Col = ColMin To ColMax
                    If Col < ColMax Then
                        RecOut = RecOut + Rec(Col, Lig) + ";"
                    Else
                        RecOut = RecOut + Rec(Col, Lig)
                    End If
                Next Col
                Print #IdxOutFile, RecOut
            Next Lig
            '===================================================================================================
            'RàZ variables
            Erase Rec()
            PageCourante = 0
            TopDebLi = True
            
            '===================================================================================================
        End If
        '===================================================================================================
    Wend
    '===================================================================================================
    Close #IndexFichier ' ferme le fichier
    Close #IdxOutFile
'===================================================================================================
End Sub
'===================================================================================================

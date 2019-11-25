Option Compare Database
'###################################################################################################
Sub LireFichierTexteParLigne()
'###################################################################################################
   
    Dim IndexFichier As Integer
    Dim MonFichier As String
    Dim ContenuLigne As String
    Dim i, k, NbPage As Integer: NbPage = 0
    Dim Top, TopTitres, TopIncPage As Boolean: Top = False: TopTitres = False: TopIncPage = False
    
    
    Dim TitresDatas() As String
    Dim Datas() As String
    
    MonFichier = "C:\Users\368790\Downloads\Extract.txt" '<-- mettez ici le nom du fichier à lire
    IndexFichier = FreeFile()
    
        
    '===================================================================================================
    Dim NumLigneSeparation As Integer: NumLigneSeparation = 0
    
    
    '===================================================================================================
    'Detecter les colonnes
    '===================================================================================================
    Open MonFichier For Input As #IndexFichier 'ouvre le fichier
    NumLigneSeparation = 0: k = 0
    '===================================================================================================
    Do While Not EOF(IndexFichier) '
        '===================================================================================================
        Line Input #IndexFichier, ContenuLigne     ' lecture du fichier ligne par ligne: la variable "ContenuLigne" contient le contenu de la ligne active
        '===================================================================================================
        'Si ligne de séparation
        If (InStr(1, ContenuLigne, " +-", vbTextCompare) > 1) Then
            If (NumLigneSeparation = 3) Then NumLigneSeparation = 0
            NumLigneSeparation = NumLigneSeparation + 1
        End If
        
        '===================================================================================================
        'Comptage NbPage
        If (TopIncPage) Then
            NbPage = NbPage + 1
        End If
        
        '===================================================================================================
        'Si ligne de titres de colonne
        If ((InStr(1, ContenuLigne, "  ! ", vbTextCompare) > 1) And (NumLigneSeparation = 1)) Then
            
            Lignes = Split(ContenuLigne, "!")
            For i = LBound(Lignes) To UBound(Lignes)
                Lignes(i) = Trim(Lignes(i))
            Next i
            
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
    Dim Rec() As String: ReDim Rec(0, Taille + 1)
    Dim RecLig, RecCol As Integer: RecLig = 0: RecCol = 0
    Open MonFichier For Input As #IndexFichier 'ouvre le fichier
    NumLigneSeparation = 0
    While Not EOF(IndexFichier) '
        '===================================================================================================
        Line Input #IndexFichier, ContenuLigne     ' lecture du fichier ligne par ligne
        '===================================================================================================
        'Si ligne de séparation
        If (InStr(1, ContenuLigne, " +-", vbTextCompare) > 1) Then
            If (NumLigneSeparation = 3) Then NumLigneSeparation = 0
            NumLigneSeparation = NumLigneSeparation + 1
        End If
        
        'Si ligne de données
        If ((InStr(1, ContenuLigne, "_! ", vbTextCompare) > 1) And (NumLigneSeparation = 2)) Then
            Lignes = Split(ContenuLigne, "!")
            For i = LBound(Lignes) To UBound(Lignes)
                Lignes(i) = Trim(Lignes(i))
            Next i
            Lignes(0) = Split(Lignes(0), "_")(0)
            'Transfert à Rec
            For i = LBound(Lignes) To UBound(Lignes)
                RecLig = RecLig + CInt(Lignes(0)) - 1
                Rec(RecLig, i) = Lignes(i)
            Next i
        End If
        

        '===================================================================================================
    Wend
    '===================================================================================================
    Close #IndexFichier ' ferme le fichier
    
    
    
'num = FreeFile
'Ouvre en écriture  et écrase un fichier précédent du même nom
'Open "C:\Users\368790\Downloads\Out.csv" For Output As #num
'Boucle sur la liste des mots
'For i = LBound(ListeMots) To UBound(ListeMots)
 'Ecrit dans le fichier texte ligne par ligne
'Print #1, ListeMots(i)
'Next i
'Fermeture
'Close #num
    
    
'===================================================================================================
End Sub
'===================================================================================================

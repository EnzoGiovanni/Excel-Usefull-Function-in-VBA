'=================================================================================================================
'Loop on access table / boucle sur chaque enregistrement
'=================================================================================================================
Sub Loop_on_access_table()
    'Varable objet "RecordSet"
    Set Table = CurrentDb.TableDefs("TableName").OpenRecordset()
    '===========================================================================================
    While Not (Table.EOF) 'Tant qu'il reste des enregistrements à traiter
    '===========================================================================================
        'Par defaut l'objet RecordSet est positionné sur le premier enregistrement
        Table.Edit '=> edition de l'enregistrement
        Trash = Table!DataName.Value 'Lecture d'une donnèe
        Table.Update '=> MàJ enregsitrement modifié

        'enregitrement suivant
        Table.MoveNext 
    '===========================================================================================
    Wend
    '===========================================================================================
    'RàZ objets
    TableName.Close: Set TableName = Nothing
End Sub

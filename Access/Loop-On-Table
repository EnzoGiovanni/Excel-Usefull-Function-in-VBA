'=================================================================================================================
'Loop on access table / boucle sur chaque enregistrement
'=================================================================================================================
Sub Loop_on_access_table()
    'Varable objet "Record Set"
    Set Table = CurrentDb.TableDefs("TableName").OpenRecordset()
    While Not (Table.EOF) 'Tant qu'il reste des enregistrements à traiter
    '------------------------------------------------------------------------------------------
        'Par defaut l'objet recordSet est positionné sur le premier enregistrement

            Table.Edit '=> edition de l'enregistrement
            
            Trash = Table!DataName.Value 'Lecture d'une donnèe
            
            Table.Update '=> MàJ enregsitrement modifié

        Table.MoveNext '=> enregitrement suivant
    '------------------------------------------------------------------------------------------
    Wend
    '------------------------------------------------------------------------------------------
    'RàZ objets
    TableName.Close: Set TableName = Nothing
End Sub

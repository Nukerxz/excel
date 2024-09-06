Sub MiseAJourStockEnBoucle()
    Dim codeBarre As String
    Dim i As Integer
    Dim action As String
    Dim reference As String
    Dim continuer As VbMsgBoxResult
    
    Do
        ' Demander à l'utilisateur de scanner le code-barre
        codeBarre = InputBox("Scannez le code-barre (ou laissez vide pour arrêter)")

        ' Si l'utilisateur appuie sur Annuler ou ne rentre rien, quitter la boucle
        If codeBarre = "" Then Exit Do
        
        ' Séparer la référence et l'action (1 ou 0)
        reference = Left(codeBarre, Len(codeBarre) - 2) ' Enlève les 2 derniers caractères (-1 ou -0)
        action = Right(codeBarre, 1) ' Récupère le dernier caractère (1 ou 0)

        ' Parcourir la colonne des références pour trouver la cartouche correspondante
        For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
            If Cells(i, 1).Value = reference Then
                If action = "1" Then
                    Cells(i, 2).Value = Cells(i, 2).Value + 1 ' Ajoute 1 au stock
                ElseIf action = "0" Then
                    Cells(i, 2).Value = Cells(i, 2).Value - 1 ' Retire 1 du stock
                End If
                Exit For
            End If
        Next i
    Loop
    
    MsgBox "Scan terminé !", vbInformation
End Sub

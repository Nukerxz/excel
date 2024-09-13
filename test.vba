Sub MiseAJourStockEnBoucle()
    Dim codeBarre As String
    Dim i As Integer
    Dim action As String
    Dim reference As String
    Dim utilisateur As String
    Dim continuer As VbMsgBoxResult
    Dim wshNetwork As Object
    Dim stock As Integer
    Dim couleur As String

    ' Créer un objet WScript.Network pour obtenir le nom de l'utilisateur Windows
    Set wshNetwork = CreateObject("WScript.Network")
    utilisateur = wshNetwork.Username

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
                ' Obtenir la valeur de stock actuelle
                stock = Cells(i, 2).Value
                
                ' Mettre à jour le stock en fonction de l'action
                If action = "1" Then
                    stock = stock + 1 ' Ajoute 1 au stock
                ElseIf action = "0" Then
                    stock = stock - 1 ' Retire 1 du stock
                End If
                
                ' Mettre à jour la cellule de stock
                Cells(i, 2).Value = stock
                
                ' Mettre à jour la colonne de la date et heure + utilisateur (colonne F)
                Cells(i, 6).Value = Now & " - " & utilisateur

                ' Vérifier la couleur dans la colonne "Couleur" (ici j'assume que la colonne "Couleur" est en colonne C)
                couleur = Cells(i, 3).Value

                ' Appliquer les règles de coloration
                If couleur = "IMAGING" Then
                    ' Si "IMAGING", colorer en rouge si stock < 2
                    If stock < 2 Then
                        Rows(i).Interior.Color = RGB(255, 0, 0) ' Colorer la ligne en rouge
                    Else
                        Rows(i).Interior.ColorIndex = xlNone ' Retirer la couleur si stock >= 2
                    End If
                Else
                    ' Si autre que "IMAGING", colorer en rouge si stock < 5
                    If stock < 5 Then
                        Rows(i).Interior.Color = RGB(255, 0, 0) ' Colorer la ligne en rouge
                    Else
                        Rows(i).Interior.ColorIndex = xlNone ' Retirer la couleur si stock >= 5
                    End If
                End If

                Exit For
            End If
        Next i
    Loop

    MsgBox "Scan terminé !", vbInformation
End Sub
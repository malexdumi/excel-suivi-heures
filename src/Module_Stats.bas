Attribute VB_Name = "Module_Stats"
' Module_Stats.bas
' Statistiques mensuelles et total cumulé depuis le début
'
' Dernier ajout : je voulais voir le total depuis que j'ai commencé
' à utiliser le fichier, pas juste semaine par semaine.
' J'ai aussi ajouté le nombre d'heures moyen par quart —
' une ligne de calcul simple mais utile.

Sub StatsMensuelles()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    ' Demander le mois
    Dim moisRef As String
    moisRef = InputBox("Quel mois ? (MM/AAAA, ex: 06/2025) :", "Stats mensuelles")
    If moisRef = "" Then Exit Sub

    Dim moisNum As Integer
    Dim anneeNum As Integer

    ' Valider le format
    If Len(moisRef) <> 7 Or Mid(moisRef, 3, 1) <> "/" Then
        MsgBox "Format attendu : MM/AAAA (ex: 06/2025)", vbExclamation
        Exit Sub
    End If

    moisNum = CInt(Left(moisRef, 2))
    anneeNum = CInt(Right(moisRef, 4))

    If moisNum < 1 Or moisNum > 12 Then
        MsgBox "Mois invalide (doit être entre 01 et 12).", vbExclamation
        Exit Sub
    End If

    ' Parcourir les transactions du mois
    Dim totalHeures As Double
    Dim totalPaie As Double
    Dim nbQuarts As Integer
    totalHeures = 0
    totalPaie = 0
    nbQuarts = 0

    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To derniereLigne
        If IsDate(ws.Cells(i, 1).Value) Then
            Dim dateLigne As Date
            dateLigne = CDate(ws.Cells(i, 1).Value)
            If Month(dateLigne) = moisNum And Year(dateLigne) = anneeNum Then
                totalHeures = totalHeures + ws.Cells(i, 4).Value
                totalPaie = totalPaie + ws.Cells(i, 5).Value
                nbQuarts = nbQuarts + 1
            End If
        End If
    Next i

    If nbQuarts = 0 Then
        MsgBox "Aucun quart trouvé pour " & moisRef & ".", vbInformation
        Exit Sub
    End If

    ' Moyenne d'heures par quart
    Dim moyenneHeures As Double
    moyenneHeures = totalHeures / nbQuarts

    MsgBox "Stats pour " & moisRef & " :" & vbNewLine & vbNewLine & _
           "Quarts travaillés  : " & nbQuarts & vbNewLine & _
           "Heures totales     : " & Format(totalHeures, "0.00") & "h" & vbNewLine & _
           "Moyenne par quart  : " & Format(moyenneHeures, "0.00") & "h" & vbNewLine & _
           "Paie estimée brute : " & Format(totalPaie, "#,##0.00") & " $", _
           vbInformation, "Stats du mois"

End Sub

' Total cumulé depuis la première entrée du fichier
Sub TotalDepuisDebut()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If derniereLigne < 2 Then
        MsgBox "Aucune donnée enregistrée.", vbInformation
        Exit Sub
    End If

    Dim totalHeures As Double
    Dim totalPaie As Double
    Dim nbQuarts As Integer
    totalHeures = 0
    totalPaie = 0
    nbQuarts = 0

    Dim i As Long
    For i = 2 To derniereLigne
        If ws.Cells(i, 4).Value <> "" Then
            totalHeures = totalHeures + ws.Cells(i, 4).Value
            totalPaie = totalPaie + ws.Cells(i, 5).Value
            nbQuarts = nbQuarts + 1
        End If
    Next i

    Dim premierDate As String
    premierDate = Format(ws.Cells(2, 1).Value, "DD/MM/YYYY")

    MsgBox "Total depuis le " & premierDate & " :" & vbNewLine & vbNewLine & _
           "Quarts enregistrés : " & nbQuarts & vbNewLine & _
           "Heures totales     : " & Format(totalHeures, "0.00") & "h" & vbNewLine & _
           "Paie totale estimée: " & Format(totalPaie, "#,##0.00") & " $", _
           vbInformation, "Total cumulé"

End Sub

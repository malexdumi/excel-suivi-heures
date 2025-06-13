Attribute VB_Name = "Module_Resume"
' Module_Resume.bas
' Résumé des heures pour une semaine donnée
'
' J'ai eu du mal avec DateDiff et Weekday() au début.
' Weekday() retourne 1 pour dimanche par défaut en VBA —
' j'ai dû utiliser vbMonday pour que la semaine commence le lundi.

Sub ResumeSemaine()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    ' Demander une date dans la semaine voulue
    Dim dateRef As String
    dateRef = InputBox("Entrer une date dans la semaine voulue (JJ/MM/AAAA) :", "Résumé semaine")
    If dateRef = "" Then Exit Sub

    If Not IsDate(dateRef) Then
        MsgBox "Date invalide.", vbExclamation
        Exit Sub
    End If

    ' Trouver le lundi de cette semaine
    Dim dateChoisie As Date
    dateChoisie = CDate(dateRef)

    ' Weekday avec vbMonday : lundi = 1, mardi = 2, ... dimanche = 7
    Dim jourDeLaSemaine As Integer
    jourDeLaSemaine = Weekday(dateChoisie, vbMonday)

    Dim lundiSemaine As Date
    lundiSemaine = dateChoisie - (jourDeLaSemaine - 1)

    Dim dimancheSemaine As Date
    dimancheSemaine = lundiSemaine + 6

    ' Parcourir les lignes et additionner les heures de cette semaine
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
        Dim dateLigne As Date
        If IsDate(ws.Cells(i, 1).Value) Then
            dateLigne = CDate(ws.Cells(i, 1).Value)
            If dateLigne >= lundiSemaine And dateLigne <= dimancheSemaine Then
                totalHeures = totalHeures + ws.Cells(i, 4).Value
                totalPaie = totalPaie + ws.Cells(i, 5).Value
                nbQuarts = nbQuarts + 1
            End If
        End If
    Next i

    ' Afficher le résumé
    Dim periode As String
    periode = Format(lundiSemaine, "DD/MM") & " au " & Format(dimancheSemaine, "DD/MM/YYYY")

    MsgBox "Semaine du " & periode & ":" & vbNewLine & vbNewLine & _
           "Nombre de quarts : " & nbQuarts & vbNewLine & _
           "Heures totales   : " & Format(totalHeures, "0.00") & "h" & vbNewLine & _
           "Paie estimée     : " & Format(totalPaie, "#,##0.00") & " $", _
           vbInformation, "Résumé de la semaine"

End Sub

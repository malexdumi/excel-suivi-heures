Attribute VB_Name = "Module_Init"
' Module_Init.bas
' Structure de la feuille et constantes globales
' -- Maria-Alexandra, juin 2025
'
' Par rapport à mon projet de budget, j'ai appris à utiliser Const
' pour les valeurs fixes — plus propre que de mettre les chiffres
' directement dans le code.

' Taux horaire (à ajuster selon le contrat)
Public Const TAUX_HORAIRE As Double = 16.25

' Nombre d'heures max par semaine avant avertissement
Public Const MAX_HEURES_SEMAINE As Integer = 40

Sub InitialiserFeuille()

    Dim ws As Worksheet
    Dim existe As Boolean
    existe = False

    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Heures" Then existe = True
    Next ws

    If Not existe Then
        ThisWorkbook.Sheets.Add.Name = "Heures"
    End If

    ' En-têtes
    With ThisWorkbook.Sheets("Heures")
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "Début"
        .Cells(1, 3).Value = "Fin"
        .Cells(1, 4).Value = "Heures travaillées"
        .Cells(1, 5).Value = "Paie estimée ($)"
        .Cells(1, 6).Value = "Note"

        ' Mettre les en-têtes en gras
        .Range(.Cells(1, 1), .Cells(1, 6)).Font.Bold = True
    End With

    MsgBox "Feuille prête !", vbInformation, "Suivi d'heures"

End Sub

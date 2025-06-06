Attribute VB_Name = "Module_Saisie"
' Module_Saisie.bas
' Ajouter un quart de travail dans la feuille
'
' Les heures en VBA sont stockées comme des fractions de 1 journée.
' Ex: 0.5 = midi, 0.75 = 18h00
' Pour convertir TimeValue("16:00") en heures : * 24

Sub AjouterQuart()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    Dim ligne As Long
    ligne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Saisie
    Dim dateQuart As String
    Dim heureDebut As String
    Dim heureFin As String
    Dim noteQuart As String

    dateQuart = InputBox("Date du quart (JJ/MM/AAAA) :", "Nouveau quart")
    If dateQuart = "" Then Exit Sub

    If Not IsDate(dateQuart) Then
        MsgBox "Date invalide. Format attendu : JJ/MM/AAAA", vbExclamation
        Exit Sub
    End If

    heureDebut = InputBox("Heure de début (ex: 16:00) :", "Nouveau quart")
    If heureDebut = "" Then Exit Sub

    heureFin = InputBox("Heure de fin (ex: 22:30) :", "Nouveau quart")
    If heureFin = "" Then Exit Sub

    noteQuart = InputBox("Note (optionnel) :", "Nouveau quart")

    ' Calculer les heures travaillées
    Dim debut As Double
    Dim fin As Double
    Dim heuresTravaillees As Double

    debut = TimeValue(heureDebut)
    fin = TimeValue(heureFin)

    ' Conversion en heures (fraction de journée × 24)
    heuresTravaillees = (fin - debut) * 24

    ' Calculer la paie estimée
    Dim paie As Double
    paie = heuresTravaillees * TAUX_HORAIRE

    ' Écrire dans la feuille
    ws.Cells(ligne, 1).Value = CDate(dateQuart)
    ws.Cells(ligne, 1).NumberFormat = "DD/MM/YYYY"
    ws.Cells(ligne, 2).Value = heureDebut
    ws.Cells(ligne, 3).Value = heureFin
    ws.Cells(ligne, 4).Value = heuresTravaillees
    ws.Cells(ligne, 4).NumberFormat = "0.00"
    ws.Cells(ligne, 5).Value = paie
    ws.Cells(ligne, 5).NumberFormat = "#,##0.00 $"
    ws.Cells(ligne, 6).Value = noteQuart

    MsgBox "Quart ajouté : " & Format(heuresTravaillees, "0.00") & "h — " & _
           Format(paie, "#,##0.00") & " $", vbInformation, "OK"

End Sub

Attribute VB_Name = "Module_Saisie"
' Module_Saisie.bas
' v1.1 — correction bug quart de nuit
'
' Bug trouvé en entrant un quart 22:00 → 00:30 :
' (00:30 - 22:00) * 24 donnait -21.5h — complètement faux.
' La raison : TimeValue("00:30") = 0.020... et TimeValue("22:00") = 0.916...
' donc fin < debut et le résultat est négatif.
'
' Fix : si le résultat est négatif, on ajoute 24h.
' C'est la solution la plus simple que j'ai trouvée.

Sub AjouterQuart()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    Dim ligne As Long
    ligne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

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

    If Not IsDate(heureDebut) Then
        MsgBox "Heure de début invalide (ex: 16:00).", vbExclamation
        Exit Sub
    End If

    heureFin = InputBox("Heure de fin (ex: 22:30) :", "Nouveau quart")
    If heureFin = "" Then Exit Sub

    If Not IsDate(heureFin) Then
        MsgBox "Heure de fin invalide (ex: 22:30).", vbExclamation
        Exit Sub
    End If

    noteQuart = InputBox("Note (optionnel) :", "Nouveau quart")

    ' Calcul des heures travaillées
    Dim debut As Double
    Dim fin As Double
    Dim heuresTravaillees As Double

    debut = TimeValue(heureDebut)
    fin = TimeValue(heureFin)
    heuresTravaillees = (fin - debut) * 24

    ' Fix quart de nuit : si négatif, le quart traverse minuit
    If heuresTravaillees < 0 Then
        heuresTravaillees = heuresTravaillees + 24
    End If

    ' Validation : un quart de plus de 14h c'est probablement une erreur de saisie
    If heuresTravaillees > 14 Then
        Dim confirmer As Integer
        confirmer = MsgBox("Ce quart fait " & Format(heuresTravaillees, "0.00") & "h — c'est correct ?", _
                           vbYesNo + vbQuestion, "Vérification")
        If confirmer = vbNo Then Exit Sub
    End If

    Dim paie As Double
    paie = heuresTravaillees * TAUX_HORAIRE

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

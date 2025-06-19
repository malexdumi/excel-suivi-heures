Attribute VB_Name = "Module_Resume"
' Module_Resume.bas
' v1.1 — avertissement si semaine dépasse MAX_HEURES_SEMAINE
'        + mise en forme de la ligne si journée longue

Sub ResumeSemaine()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    Dim dateRef As String
    dateRef = InputBox("Entrer une date dans la semaine voulue (JJ/MM/AAAA) :", "Résumé semaine")
    If dateRef = "" Then Exit Sub

    If Not IsDate(dateRef) Then
        MsgBox "Date invalide.", vbExclamation
        Exit Sub
    End If

    Dim dateChoisie As Date
    dateChoisie = CDate(dateRef)

    Dim jourDeLaSemaine As Integer
    jourDeLaSemaine = Weekday(dateChoisie, vbMonday)

    Dim lundiSemaine As Date
    lundiSemaine = dateChoisie - (jourDeLaSemaine - 1)

    Dim dimancheSemaine As Date
    dimancheSemaine = lundiSemaine + 6

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
            If dateLigne >= lundiSemaine And dateLigne <= dimancheSemaine Then
                totalHeures = totalHeures + ws.Cells(i, 4).Value
                totalPaie = totalPaie + ws.Cells(i, 5).Value
                nbQuarts = nbQuarts + 1
            End If
        End If
    Next i

    Dim periode As String
    periode = Format(lundiSemaine, "DD/MM") & " au " & Format(dimancheSemaine, "DD/MM/YYYY")

    ' Avertissement si on dépasse le maximum de la semaine
    ' (utilise la constante définie dans Module_Init)
    If totalHeures > MAX_HEURES_SEMAINE Then
        MsgBox "⚠ Attention : " & Format(totalHeures, "0.00") & "h cette semaine — " & _
               "tu dépasses les " & MAX_HEURES_SEMAINE & "h !", _
               vbExclamation, "Semaine chargée"
    End If

    MsgBox "Semaine du " & periode & ":" & vbNewLine & vbNewLine & _
           "Nombre de quarts : " & nbQuarts & vbNewLine & _
           "Heures totales   : " & Format(totalHeures, "0.00") & "h" & vbNewLine & _
           "Paie estimée     : " & Format(totalPaie, "#,##0.00") & " $", _
           vbInformation, "Résumé de la semaine"

End Sub

' Colore en orange les lignes où on a travaillé plus de 8h d'affilée
' -- idée ajoutée après coup, pratique pour repérer les grosses journées
Sub MarquerLonguesJournees()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Heures")

    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To derniereLigne
        Dim heures As Double
        heures = ws.Cells(i, 4).Value

        If heures > 8 Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).Interior.Color = RGB(255, 220, 150)
        Else
            ' Remettre blanc si on a corrigé une entrée
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).Interior.ColorIndex = xlNone
        End If
    Next i

End Sub

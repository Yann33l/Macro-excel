'
' Made by YLE
'
Option Explicit

Sub NetoyageDesDonnées(nomFeuille As String)

    Dim s As Integer
    Application.ScreenUpdating = False
    On Error GoTo CleanExit
    
    For s = 1 To ThisWorkbook.Sheets.Count - 2
        Worksheets(nomFeuille & s).Activate
        If nomFeuille = "ADN Maxwell custom " Then
            Range("C20:K35,D7,D10:D11,D14:D16").Select
            Selection.ClearContents
            Range("B20:K35").Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        ElseIf nomFeuille = "ARN Maxwell " Then
            Range("C20:L35,D7,D10:D11,D14:D16").Select
            Selection.ClearContents
            Range("B20:L35").Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
        End If
        Worksheets("ExportAriane").Activate
    Next s

CleanExit:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Erreur: " & Err.Description
End Sub

Function ValeurSaisiEstElleUneDate(question As String, titre As String, Default As String) As Date
Dim reponseStr As String
Dim reponseDate As Date
    Do
        reponseStr = InputBox(question, titre, Default)
        If IsDate(reponseStr) Then
            reponseDate = CDate(reponseStr)
            ValeurSaisiEstElleUneDate = reponseDate
            Exit Do
        Else
            MsgBox "Date au format JJ/MM/AAAA obligatoire"
        End If
    Loop

End Function

Sub CreationSeriesTechniques()

    Dim exportAriane As Worksheet
    Dim ws As Worksheet
    Dim numeroFormulaire As String
    Dim typeDeSerieExtraction As String
    Dim nomFeuille As String
    Dim numPremiereSerie As Integer
    Dim numPremierBlanc As Integer
    Dim dateJ1 As Date
    Dim dateJ2 As Date
    Dim opérateursJ1 As String
    Dim opérateursJ2 As String
    Dim volumeElution As Integer
    Dim nbrSerieACreer As Integer
    Dim numBlanc As String
    Dim numSerie As Integer
    Dim nomSerie As String
    Dim nbrLigneDeLaSerie As Integer
    Dim positionAléatoireBlanc As Integer
    Dim s As Integer
    Dim ligne As Integer
    Dim LignePat As Range
             
    ThisWorkbook.Sheets("ExportAriane").Activate
    Set exportAriane = ActiveSheet
    
    numeroFormulaire = Left(ThisWorkbook.Name, 11)
    If numeroFormulaire = "PAM-FQ-0162" Then
        nomFeuille = "ADN Maxwell custom "
        typeDeSerieExtraction = "EXTR.ADN.FIXE"
        volumeElution = 70
    ElseIf numeroFormulaire = "PAM-FQ-0206" Then
        nomFeuille = "ARN Maxwell "
        typeDeSerieExtraction = "EXTR.ARN.FIXE"
        volumeElution = 50
    End If
    
    numPremiereSerie = Application.InputBox("n° de la 1ere série : (ex: 4625)", "numSerie", Type:=1)
    numPremierBlanc = Application.InputBox("Quel est le nom du premier blanc (ex pour M154 saisir: 154) :", "premierBlanc", Type:=1)
    dateJ1 = ValeurSaisiEstElleUneDate("Saisir la Date J1 au format jj/mm/aaaa", "dateJ1", Date)
    dateJ2 = ValeurSaisiEstElleUneDate("Date J2 :", "dateJ2", dateJ1 + 1)
    opérateursJ1 = Application.InputBox("opérateur(s) J1 (ex: ABA2/LDU1) :", "opé J1", Type:=2)
    opérateursJ2 = Application.InputBox("opérateur(s) J2 (ex: ABA2/LDU1/VNO):", "opé J2", opérateursJ1, Type:=2)
    volumeElution = Application.InputBox("Volume élution (µL):", "Ve", volumeElution, Type:=1)
    
    NetoyageDesDonnées (nomFeuille)
    
    If Range("B62") <> "" Then
        nbrSerieACreer = 5
    ElseIf Range("B47") <> "" Then
        nbrSerieACreer = 4
    ElseIf Range("B32") <> "" Then
        nbrSerieACreer = 3
    ElseIf Range("B17") <> "" Then
        nbrSerieACreer = 2
    Else: nbrSerieACreer = 1
    End If

    For s = 1 To nbrSerieACreer
        Set ws = Worksheets(nomFeuille & s)
        numBlanc = numPremierBlanc + s - 1
        numSerie = Right(numPremiereSerie, 4) + s - 1
        nomSerie = "ST-" & Right(Year(dateJ1), 2) & "-" & typeDeSerieExtraction & "-" & Format(numSerie, "0000")
        nbrLigneDeLaSerie = Application.WorksheetFunction.CountA(exportAriane.Range(exportAriane.Cells(2 - 15 + s * 15, "B"), exportAriane.Cells(16 - 15 + s * 15, "B")))
        positionAléatoireBlanc = (nbrLigneDeLaSerie * Rnd) + 1

        
        With ws
            .Range("D7").Value = nomSerie
            .Range("D10").Value = dateJ1
            .Range("D11").Value = opérateursJ1
            .Range("D14").Value = dateJ2
            .Range("D15").Value = opérateursJ2
            .Range("D16").Value = volumeElution & " µL"
        End With
        
        For ligne = 1 To nbrLigneDeLaSerie + 1
            If ligne < positionAléatoireBlanc Then
                Set LignePat = exportAriane.Range(exportAriane.Cells(1 + ligne + (s - 1) * 15, "B"), exportAriane.Cells(1 + ligne + (s - 1) * 15, "F"))
                ws.Range(ws.Cells(19 + ligne, "C"), ws.Cells(19 + ligne, "G")).Value = LignePat.Value
            ElseIf ligne = positionAléatoireBlanc Then
                ws.Cells(19 + ligne, "C").Value = "BLANC M" & Format(numBlanc, "000")
                    If nomFeuille = "ADN Maxwell custom " Then
                        ws.Range(ws.Cells(19 + ligne, "B"), ws.Cells(19 + ligne, "K")).Interior.Color = 49407
                    ElseIf nomFeuille = "ARN Maxwell " Then
                        ws.Range(ws.Cells(19 + ligne, "B"), ws.Cells(19 + ligne, "L")).Interior.Color = 49407
                    End If
            Else
                Set LignePat = exportAriane.Range(exportAriane.Cells(1 + ligne - 1 + (s - 1) * 15, "B"), exportAriane.Cells(1 + ligne - 1 + (s - 1) * 15, "F"))
                ws.Range(ws.Cells(19 + ligne, "C"), ws.Cells(19 + ligne, "G")).Value = LignePat.Value
            End If
        Next ligne
    Next s
    
    Worksheets(nomFeuille & "1").Activate
                      
End Sub





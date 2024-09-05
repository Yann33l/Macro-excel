'
' Made by YLE
'
Function ExtraireEntre(texte As String, charDebut As String, charFin As String) As String

    Dim StartPosition As Long
    Dim EndPosition As Long
    
    positionDebut = InStr(texte, charDebut)
    positionFin = InStr(texte, charFin)
    
    If positionDebut = 0 And positionFin = 0 Then
        ExtraireEntre = ""
    Else
        ExtraireEntre = Mid(texte, positionDebut + 1, positionFin - positionDebut - 1)
    End If

End Function

Function ExtraireJusquA(texte As String, charFin As String) As String

    Dim positionFin As Long
    
    positionFin = InStr(texte, charFin)
    If positionFin = 0 Then
        ExtraireJusquA = texte
    Else
        ExtraireJusquA = Left(texte, positionFin - 1)
    End If
    
End Function

Function CreerDossiersSiNonExistants(PathFeuilleDeComptageInitial As String, anneeSt As Integer, numSerie As String)

    Dim cheminTest As String
    Dim cheminSerie As String
    
    cheminTest = PathFeuilleDeComptageInitial & anneeSt
    If Dir(cheminTest, vbDirectory) = "" Then
        MkDir cheminTest
    End If

    cheminSerie = cheminTest & Application.PathSeparator & "SERIE " & numSerie
    If Dir(cheminSerie, vbDirectory) = "" Then
        MkDir cheminSerie
    End If

    CreerDossiersSiNonExistants = cheminSerie
    
End Function


Function ValeurSaisiEstElleUneDate(question As String, titre As String, Default As String)

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

Function LaFeuilleEstElleVierge() As Boolean
    If IsEmpty(Range("C7,C8,G7,I7,L7,P7,P8").Value) Then
        LaFeuilleEstElleVierge = True
    Else
        LaFeuilleEstElleVierge = False
    End If
End Function

Sub CreationFeuilleDeComptage()

    Dim feuilleDeSerie As Worksheet
    Dim feuilleDeComptage As Worksheet
    Dim nbrLignes As Long
    Dim numSerie As String
    Dim technicien As String
    Dim numPat As String
    Dim nomPat As String
    Dim dateDemande As Date
    Dim nomFeuilleDeComptage As String
    Dim datePremiereLecture As Date
    Dim dateTechnique As Date
    Dim anneeSt As Integer
    Dim PathFeuilleDeComptageInitial As String
    Dim cheminEnregistrement As String
    Dim cheminSerie As String
    Dim numHisto As String
    Dim numLame As String
    
    
    Set feuilleDeSerie = ThisWorkbook.ActiveSheet
    feuilleDeSerie.Activate
    ActiveWindow.ActivateNext
    If ActiveWorkbook.Name Like "PAM-FQ-0027*" Or ActiveWorkbook.Name Like "PAM-FQ-0110*" Then
        If LaFeuilleEstElleVierge() = True Then
            Set feuilleDeComptage = ActiveSheet
        Else
            MsgBox "La feuille de comptage PAM-FQ-0027 ou PAM-FQ-0110 doit être vierge."
            Exit Sub
        End If
    Else
        MsgBox "Merci d'ouvrir la feuille de comptage PAM-FQ-0027 ou PAM-FQ-0110."
        Exit Sub
    End If
        
    feuilleDeSerie.Activate
    If IsDate(Range("C11").Value) Then
        dateTechnique = Range("C11")
    Else
        Range("C11") = ValeurSaisiEstElleUneDate("Date de technique : ", "Date technique", Date)
        dateTechnique = Range("C11")
    End If
        
    
    feuilleDeComptage.Activate
    PathFeuilleDeComptageInitial = ActiveWorkbook.Path & Application.PathSeparator
    nomFeuilleDeComptage = Replace(ActiveWindow.Caption, ".xlsx", "")
    Range("F10:J10").Value = "Visa lecteur pathologiste/ingénieur: " & InputBox("Visa second lecteur : ", "secondLecteur", "CLA1")
    datePremiereLecture = ValeurSaisiEstElleUneDate("Date de première lecture prévu : ", "DateLecture", dateTechnique + 1)
    Range("P7").Value = datePremiereLecture
        
    feuilleDeSerie.Activate
    nbrLignes = feuilleDeSerie.Evaluate("COUNTA(B17:B28)")
    numSerie = Right(feuilleDeSerie.Range("C9").Value, 4)
    technicien = Range("C12").Value
    anneeSt = Year(dateTechnique)
    
    cheminEnregistrement = CreerDossiersSiNonExistants(PathFeuilleDeComptageInitial, anneeSt, numSerie)

    For i = 1 To nbrLignes
        numPat = Cells(16 + i, "B").Value
        nomPat = Cells(16 + i, "C").Value
        dateDemande = Cells(16 + i, "H").Value
        numHisto = ExtraireJusquA(numPat, " ")
        numLame = ExtraireEntre(numPat, "(", ")")
        If numLame = "" Then
            codeLame = ""
        Else
            codeLame = Right(numLame, Len(numLame) - Len("LAME "))
        End If
        
        If codeLame Like "*[a-zA-Z]*" Then
            Lame = " (" & codeLame & ")"
        Else
            Lame = ""
        End If

        feuilleDeComptage.Activate
        
        Range("I7").Value = technicien
        Range("G7").Value = numSerie
        Range("C7").Value = numHisto
        Range("C8").Value = nomPat
        Range("L7").Value = dateDemande
        Range("Q13:Q15") = numLame
        Range("A10:E10").Value = "Visa lecteur technicien: " & technicien

        ActiveWorkbook.SaveAs Filename:= _
            cheminEnregistrement & Application.PathSeparator & nomFeuilleDeComptage & " " _
            & numHisto & Lame & ".xlsx", _
            FileFormat:=xlOpenXMLWorkbook
        feuilleDeSerie.Activate
    Next
    feuilleDeComptage.Activate
    ActiveWorkbook.Close
    MsgBox "Les feuilles de comptages ont été créées."
End Sub

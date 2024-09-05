'
' Made by YLE
'

Function tempsPepsine(Sonde As String) As String
    Select Case Sonde
        Case "HER2"
            tempsPepsine = "3'"
        Case "Sarcome"
            tempsPepsine = "7'"
        Case "ALK_BA"
            tempsPepsine = "5'30"
    End Select
End Function

Sub CreationFeuilleDeTravail()

Dim feuilleDeSerie As Worksheet
Dim feuilleDeTravail As Worksheet
Dim dateTechnique As Date
Dim patients As Range
Dim demandeurs As Range
Dim sondes As Range
Dim operateur As String
Dim numSerie As Integer
Dim nbrPatient As Integer
Dim coche As String
Dim isUrgent As Boolean
Dim fixateur As String

Set feuilleDeSerie = ThisWorkbook.ActiveSheet
feuilleDeSerie.Activate

ActiveWindow.ActivateNext
If ActiveWorkbook.Name Like "PAM-FQ-0030*" Then
    Set feuilleDeTravail = ActiveSheet
Else
    MsgBox "La feuille de travail PAM-FQ-0030 doit être ouverte."
    Exit Sub
End If
        
feuilleDeSerie.Activate
dateTechnique = Range("C11")
Set patients = Range("B17:C28")
Set demandeurs = Range("D17:D28")
Set sondes = Range("E17:E28")
operateur = Range("C12").Value
numSerie = Right(Range("C9").Value, 4)
nbrPatient = feuilleDeSerie.Evaluate("COUNTA(B17:B28)")

coche = "X"

ActiveWindow.ActivateNext
feuilleDeTravail.Activate

Range("D6:D8,C11:M22,O11:Q22").Select
Selection.ClearContents

Range("C11:D22").Value = patients.Value
Range("P11:P22").Value = demandeurs.Value
Range("L11:L22").Value = sondes.Value
Range("D6").Value = dateTechnique
Range("D7").Value = operateur
Range("D8").Value = numSerie

For p = 1 To nbrPatient
    feuilleDeSerie.Activate
    isUrgent = Cells(16 + p, "A")
    fixateur = Cells(16 + p, "G")
    
    feuilleDeTravail.Activate
    Cells(10 + p, "E").Value = coche
    
    If fixateur = "Formol" Then
        Cells(10 + p, "I").Value = coche
    Else: Cells(10 + p, "K").Value = coche
    End If
    
    If Cells(10 + p, "L") = "FISH.HER2-SEIN" Or Cells(10 + p, "L") = "FISH.HER2-HS" Or Cells(10 + p, "L") Like "FISH.AMP*" Then
        Cells(10 + p, "M").Value = "AMP"
    Else: Cells(10 + p, "M").Value = "BA"
    End If
    
    If Cells(10 + p, "L") = "FISH.HER2-SEIN" Or Cells(10 + p, "L") = "FISH.HER2-HS" Then
        Cells(10 + p, "O").Value = tempsPepsine("HER2")
    ElseIf Cells(10 + p, "L") = "FISH.ALK-BA" Or Cells(10 + p, "L") = "FISH.ALK-BA.POU" Or Cells(10 + p, "L") = "FISH.ALK-BA.AUT" Then
        Cells(10 + p, "O").Value = tempsPepsine("ALK_BA")
    Else: Cells(10 + p, "O").Value = tempsPepsine("Sarcome")
    End If
    
    If isUrgent = True Then
        With Cells(10 + p, "Q")
            .Value = "Urgent"
            .Font.Size = 16
            .Font.Color = vbRed
            .Font.Bold = True
        End With
    Else
        With Cells(10 + p, "Q")
            .Font.Size = 10
            .Font.Color = vbBlack
            .Font.Bold = False
        End With
    End If
Next

End Sub

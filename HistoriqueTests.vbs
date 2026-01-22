Option Explicit

' --- Configuration et Initialisation ---
Dim objFSO, objFolder, logfile, objFile, line, tmp, sn, fOut
Dim workingPath, resultFile, results, dictGlobalHistory
Dim fileCount, currentSN, currentStatus

Set objFSO = CreateObject("Scripting.FileSystemObject")
workingPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(workingPath)

' Constantes pour l'encodage (0 = ANSI)
Const TristateFalse = 0 
Const MARKER_OK = "[PROD_OK]"
Const MARKER_ERROR = "[PROD_ERROR]"

resultFile = workingPath & "\Historique_Tests_Complet.txt"
fileCount = 0

' Initialisation du dictionnaire global
Set dictGlobalHistory = CreateObject("Scripting.Dictionary")

' --- 1. Parcours des fichiers de logs ---
For Each logfile In objFolder.Files
    ' Vérification du nom de fichier (insensible à la casse)
    If InStr(1, logfile.Name, "elisa_prod_log", vbTextCompare) = 1 Then
        fileCount = fileCount + 1
        
        ' Ouverture forcée en mode ANSI
        Set objFile = objFSO.OpenTextFile(logfile.Path, 1, False, TristateFalse)
        
        currentSN = ""
        currentStatus = "Test Incomplet"
        
        Do Until objFile.AtEndOfStream
            line = objFile.ReadLine
            line = Replace(line, Chr(0), "") ' Nettoyage des octets nuls
            
            ' Recherche insensible à la casse de Datamatrix
            If InStr(1, line, "Datamatrix", vbTextCompare) > 0 Then
                ' Enregistrement du passage précédent
                If currentSN <> "" Then
                    AddPassage dictGlobalHistory, currentSN, currentStatus, logfile.Name
                End If
                
                ' Nouveau bloc détecté
                tmp = Split(line, "#")
                If UBound(tmp) >= 2 Then
                    currentSN = CleanSerial(tmp(2))
                    currentStatus = "Test Incomplet"
                    UpdateStatus line, currentStatus
                Else
                    currentSN = ""
                End If
            ElseIf currentSN <> "" Then
                UpdateStatus line, currentStatus
            End If
        Loop
        
        ' Enregistrement du dernier bloc
        If currentSN <> "" Then
            AddPassage dictGlobalHistory, currentSN, currentStatus, logfile.Name
        End If
        
        objFile.Close
    End If
Next

' --- 2. Construction du rapport final ---
results = "HISTORIQUE COMPLET DES PASSAGES PAR NUMERO DE SERIE" & vbCrLf & _
          "Genere le : " & Now() & vbCrLf & _
          "Fichiers analyses : " & fileCount & vbCrLf & _
          String(80, "=") & vbCrLf & vbCrLf

If dictGlobalHistory.Count = 0 Then
    results = results & "Aucun numero de serie trouve dans les logs." & vbCrLf
Else
    Dim allSN, k
    allSN = dictGlobalHistory.Keys
    SortArray allSN

    For Each k In allSN
        results = results & "SN: " & k & vbCrLf
        results = results & dictGlobalHistory(k) & vbCrLf
        results = results & String(60, "-") & vbCrLf
    Next
End If

' --- 3. Ecriture du fichier de sortie ---
On Error Resume Next
' Retour à False (ANSI) pour éviter l'effet "espaces entre les lettres"
Set fOut = objFSO.CreateTextFile(resultFile, True, False) 
If Err.Number = 0 Then
    fOut.Write results
    fOut.Close
Else
    WScript.Echo "ERREUR : Impossible de créer le fichier d'historique (Erreur " & Err.Number & ")."
End If
On Error GoTo 0

' --- Résumé Final ---
WScript.Echo "========================================" & vbCrLf & _
             "Historique des tests termine." & vbCrLf & _
             "Passages analyses : " & fileCount & " fichiers logs" & vbCrLf & _
             "Numeros de serie   : " & dictGlobalHistory.Count & vbCrLf & _
             "Fichier genere     : Historique_Tests_Complet.txt" & vbCrLf & _
             "========================================"

' --- Fonctions et Sous-programmes ---

Sub UpdateStatus(ByVal currentLine, ByRef status)
    If InStr(1, currentLine, MARKER_OK, vbTextCompare) > 0 Then
        status = "Test OK"
    ElseIf InStr(1, currentLine, MARKER_ERROR, vbTextCompare) > 0 Then
        ' On nettoie la ligne d'erreur pour ne pas faire planter l'écriture ANSI plus tard
        status = Sanitize(Trim(currentLine))
    End If
End Sub

Function Sanitize(str)
    Dim i, c, res
    res = ""
    For i = 1 To Len(str)
        c = AscW(Mid(str, i, 1))
        ' On ne garde que les caractères ANSI standards (0-255)
        ' Si un caractère est au-delà, on le remplace par un espace pour éviter le crash
        If c >= 0 And c <= 255 Then
            res = res & Mid(str, i, 1)
        Else
            res = res & " "
        End If
    Next
    Sanitize = res
End Function

Sub AddPassage(dict, sn, status, fileName)
    Dim passageInfo
    passageInfo = "   -> " & Left(status & Space(50), 50) & " | Fichier: " & fileName & vbCrLf
    If Not dict.Exists(sn) Then
        dict.Add sn, passageInfo
    Else
        dict(sn) = dict(sn) & passageInfo
    End If
End Sub

Function SortArray(arr)
    Dim i, j, temp
    If Not IsArray(arr) Then Exit Function
    If UBound(arr) < 1 Then Exit Function
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrCompare(arr(i), arr(j)) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next
    Next
End Function

Function StrCompare(a, b)
    StrCompare = StrComp(a, b, vbTextCompare)
End Function

Function CleanSerial(value)
    Dim cleaned
    cleaned = Trim(value)
    cleaned = Replace(cleaned, Chr(0), "")
    cleaned = Replace(cleaned, "=", "")
    cleaned = Replace(cleaned, vbTab, "")
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    CleanSerial = Trim(cleaned) ' Trim final après remplacement
End Function

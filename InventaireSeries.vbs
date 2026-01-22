Option Explicit

' --- Configuration et Initialisation ---
Dim objFSO, objFolder, logfile, objFile, line, tmp, sn, fOut
Dim workingPath, resultFile, results, dictSerialsInFile
Dim fileCount, totalSNCount

Set objFSO = CreateObject("Scripting.FileSystemObject")
workingPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(workingPath)

' Constantes pour l'encodage (0 = ANSI selon LLM_Instructions.md)
Const TristateFalse = 0 

resultFile = workingPath & "\Inventaire_Series.txt"
results = "INVENTAIRE DES NUMEROS DE SERIE PAR FICHIER LOG" & vbCrLf & _
          "Genere le : " & Now() & vbCrLf & _
          String(60, "=") & vbCrLf & vbCrLf

fileCount = 0
totalSNCount = 0

' --- 1. Parcours des fichiers de logs ---
For Each logfile In objFolder.Files
    If InStr(1, logfile.Name, "elisa_prod_log", vbTextCompare) = 1 Then
        fileCount = fileCount + 1
        ' Utilisation d'un dictionnaire local pour éviter les doublons dans un même fichier log
        Set dictSerialsInFile = CreateObject("Scripting.Dictionary")
        
        Set objFile = objFSO.OpenTextFile(logfile.Path, 1, False, TristateFalse)
        
        Do Until objFile.AtEndOfStream
            line = objFile.ReadLine
            line = Replace(line, Chr(0), "") ' Nettoyage des octets nuls (LLM_Instructions)
            
            If InStr(line, "Datamatrix") > 0 Then
                tmp = Split(line, "#")
                If UBound(tmp) >= 2 Then
                    sn = CleanSerial(tmp(2))
                    If sn <> "" Then
                        If Not dictSerialsInFile.Exists(sn) Then
                            dictSerialsInFile.Add sn, True
                        End If
                    End If
                End If
            End If
        Loop
        objFile.Close
        
        ' Ajout des résultats du fichier actuel
        results = results & "FICHIER : " & logfile.Name & vbCrLf
        results = results & "Nombre de SN uniques : " & dictSerialsInFile.Count & vbCrLf
        results = results & "Liste : " & Join(dictSerialsInFile.Keys, ", ") & vbCrLf
        results = results & String(60, "-") & vbCrLf & vbCrLf
        
        totalSNCount = totalSNCount + dictSerialsInFile.Count
        Set dictSerialsInFile = Nothing
    End If
Next

' --- 2. Ecriture du fichier de sortie ---
On Error Resume Next
Set fOut = objFSO.CreateTextFile(resultFile, True, False) ' False = ANSI
If Err.Number = 0 Then
    fOut.Write results
    fOut.Close
Else
    WScript.Echo "ERREUR : Impossible de créer le fichier d'inventaire."
End If
On Error GoTo 0

' --- Résumé Final (Unique selon LLM_Instructions) ---
WScript.Echo "========================================" & vbCrLf & _
             "Inventaire des logs termine." & vbCrLf & _
             "Fichiers analyses      : " & fileCount & vbCrLf & _
             "Fichier genere         : Inventaire_Series.txt" & vbCrLf & _
             "========================================"

' --- Fonctions Utilitaires (Normalisation selon LLM_Instructions) ---
Function CleanSerial(value)
    Dim cleaned
    cleaned = Trim(value)
    cleaned = Replace(cleaned, Chr(0), "")
    cleaned = Replace(cleaned, "=", "")
    cleaned = Replace(cleaned, vbTab, "")
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    CleanSerial = cleaned
End Function

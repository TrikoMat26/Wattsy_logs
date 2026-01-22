Option Explicit

' --- Configuration et Initialisation ---
Dim objFSO, objFolder, dictSerials, dictResults, dictCounts
Dim logfile, objFile, currentSerial, line, tmp, sn, f
Dim workingPath, serialFile, serial, i

Set objFSO = CreateObject("Scripting.FileSystemObject")
workingPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(workingPath)

' Constante pour l'encodage (0 = ANSI)
Const TristateFalse = 0 

' Utilisation de dictionnaires pour la performance
Set dictSerials = CreateObject("Scripting.Dictionary")
Set dictResults = CreateObject("Scripting.Dictionary")
Set dictCounts = CreateObject("Scripting.Dictionary")

' --- 1. Lecture des numéros de série cibles ---
Dim targetFile : targetFile = workingPath & "\NumSerieKO.txt"
If objFSO.FileExists(targetFile) Then
    Set serialFile = objFSO.OpenTextFile(targetFile, 1, False, TristateFalse)
    Do Until serialFile.AtEndOfStream
        serial = Trim(serialFile.ReadLine)
        serial = Replace(serial, Chr(0), "")
        If serial <> "" Then
            If Not dictSerials.Exists(serial) Then
                dictSerials.Add serial, True
                dictResults.Add serial, ""
                dictCounts.Add serial, 0
            End If
        End If
    Loop
    serialFile.Close
Else
    WScript.Echo "ERREUR : Fichier NumSerieKO.txt introuvable."
    WScript.Quit 1
End If

' --- 2. Parcours des fichiers de logs (Un seul passage) ---
For Each logfile In objFolder.Files
    If InStr(1, logfile.Name, "elisa_prod_log", vbTextCompare) = 1 Then
        Set objFile = objFSO.OpenTextFile(logfile.Path, 1, False, TristateFalse)
        currentSerial = ""
        
        Do Until objFile.AtEndOfStream
            line = objFile.ReadLine
            line = Replace(line, Chr(0), "")
            
            If InStr(line, "Datamatrix") > 0 Then
                tmp = Split(line, "#")
                If UBound(tmp) >= 2 Then
                    sn = CleanSerial(tmp(2))
                    If dictSerials.Exists(sn) Then
                        currentSerial = sn
                        dictCounts(sn) = dictCounts(sn) + 1
                        
                        ' Ajout d'un saut de ligne si ce n'est pas le début du dictionnaire pour ce numéro
                        If dictResults(sn) <> "" Then dictResults(sn) = dictResults(sn) & vbCrLf
                        
                        ' Ajout de l'en-tête d'occurrence systématique (Fichier + Numéro)
                        dictResults(sn) = dictResults(sn) & String(40, "-") & vbCrLf & _
                                          "Occurrence #" & dictCounts(sn) & " (Fichier: " & logfile.Name & ")" & vbCrLf & _
                                          String(40, "-") & vbCrLf
                        
                        dictResults(sn) = dictResults(sn) & line
                    Else
                        currentSerial = "" ' Datamatrix d'un numéro non suivi : on arrête la collecte
                    End If
                Else
                    currentSerial = "" ' Datamatrix mal formé : on arrête la collecte
                End If
            ElseIf currentSerial <> "" Then
                dictResults(currentSerial) = dictResults(currentSerial) & vbCrLf & line
            End If
        Loop
        objFile.Close
    End If
Next

' --- 3. Ecriture des fichiers de sortie ---
Dim keys, k, count
keys = dictResults.Keys
count = 0

For Each k In keys
    If dictResults(k) <> "" Then
        On Error Resume Next
        Set f = objFSO.CreateTextFile(workingPath & "\" & k & ".txt", True, False)
        If Err.Number = 0 Then
            f.Write dictResults(k)
            f.Close
            count = count + 1
        End If
        On Error GoTo 0
    End If
Next

' --- Résumé Final ---
WScript.Echo "========================================" & vbCrLf & _
             "Analyse terminee." & vbCrLf & _
             "Numeros recherches : " & dictSerials.Count & vbCrLf & _
             "Fichiers generes   : " & count & vbCrLf & _
             "========================================"

' --- Fonctions Utilitaires ---
Function CleanSerial(value)
    Dim cleaned
    cleaned = Trim(value)
    cleaned = Replace(cleaned, Chr(0), "") ' Supprime les octets nuls
    cleaned = Replace(cleaned, "=", "")
    cleaned = Replace(cleaned, vbTab, "")
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    CleanSerial = cleaned
End Function
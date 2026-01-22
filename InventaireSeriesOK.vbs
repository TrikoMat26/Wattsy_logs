Option Explicit

' --- Configuration et Initialisation ---
Dim objFSO, objFolder, logfile, objFile, line, tmp, sn, fOut
Dim workingPath, resultFile, results, dictSerialsOKInFile
Dim fileCount, currentSN, isBlockOK

Set objFSO = CreateObject("Scripting.FileSystemObject")
workingPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(workingPath)

' Constantes pour l'encodage (0 = ANSI selon LLM_Instructions.md)
Const TristateFalse = 0 
Const OK_CRITERIA = "[PROD_OK]: Calibration and test passed"

resultFile = workingPath & "\Inventaire_Series_OK.txt"
results = "INVENTAIRE DES NUMEROS DE SERIE AVEC TEST OK (SUCCESSFUL) PAR FICHIER LOG" & vbCrLf & _
          "Genere le : " & Now() & vbCrLf & _
          "Critere : Contient " & OK_CRITERIA & vbCrLf & _
          String(80, "=") & vbCrLf & vbCrLf

fileCount = 0

' --- 1. Parcours des fichiers de logs (Un seul passage) ---
Dim dictGlobalOverview : Set dictGlobalOverview = CreateObject("Scripting.Dictionary")
Dim sameFileDuplicates : sameFileDuplicates = ""
Dim crossFileDuplicates : crossFileDuplicates = ""

For Each logfile In objFolder.Files
    If InStr(1, logfile.Name, "elisa_prod_log", vbTextCompare) = 1 Then
        fileCount = fileCount + 1
        Set dictSerialsOKInFile = CreateObject("Scripting.Dictionary")
        
        Set objFile = objFSO.OpenTextFile(logfile.Path, 1, False, TristateFalse)
        currentSN = ""
        isBlockOK = False
        
        Do Until objFile.AtEndOfStream
            line = objFile.ReadLine
            line = Replace(line, Chr(0), "")
            
            If InStr(line, "Datamatrix") > 0 Then
                If currentSN <> "" And isBlockOK Then
                    dictSerialsOKInFile(currentSN) = dictSerialsOKInFile(currentSN) + 1
                    If Not dictGlobalOverview.Exists(currentSN) Then dictGlobalOverview.Add currentSN, CreateObject("Scripting.Dictionary")
                    dictGlobalOverview(currentSN)(logfile.Name) = dictGlobalOverview(currentSN)(logfile.Name) + 1
                End If
                
                tmp = Split(line, "#")
                If UBound(tmp) >= 2 Then
                    currentSN = CleanSerial(tmp(2))
                    isBlockOK = (InStr(line, OK_CRITERIA) > 0)
                Else
                    currentSN = ""
                End If
            ElseIf currentSN <> "" Then
                If Not isBlockOK Then
                    If InStr(line, OK_CRITERIA) > 0 Then isBlockOK = True
                End If
            End If
        Loop
        
        If currentSN <> "" And isBlockOK Then
            dictSerialsOKInFile(currentSN) = dictSerialsOKInFile(currentSN) + 1
            If Not dictGlobalOverview.Exists(currentSN) Then dictGlobalOverview.Add currentSN, CreateObject("Scripting.Dictionary")
            dictGlobalOverview(currentSN)(logfile.Name) = dictGlobalOverview(currentSN)(logfile.Name) + 1
        End If
        objFile.Close
        
        If dictSerialsOKInFile.Count > 0 Then
            results = results & "FICHIER : " & logfile.Name & vbCrLf
            results = results & "Nombre de SN OK (distincts) : " & dictSerialsOKInFile.Count & vbCrLf
            
            Dim keysArray, k, first, subDup
            subDup = ""
            keysArray = dictSerialsOKInFile.Keys
            
            results = results & " - Par ordre d'apparition : "
            first = True
            For Each k In keysArray
                If Not first Then results = results & ", "
                results = results & k
                If dictSerialsOKInFile(k) > 1 Then 
                    subDup = subDup & "   - " & k & " (" & dictSerialsOKInFile(k) & " fois)" & vbCrLf
                End If
                first = False
            Next
            results = results & vbCrLf
            
            results = results & " - Par ordre croissant   : "
            SortArray keysArray 
            first = True
            For Each k In keysArray
                If Not first Then results = results & ", "
                results = results & k
                first = False
            Next
            results = results & vbCrLf
            
            If subDup <> "" Then
                sameFileDuplicates = sameFileDuplicates & "Dans " & logfile.Name & " :" & vbCrLf & subDup & vbCrLf
            End If
            results = results & String(80, "-") & vbCrLf & vbCrLf
        End If
        Set dictSerialsOKInFile = Nothing
    End If
Next

For Each currentSN In dictGlobalOverview.Keys
    If dictGlobalOverview(currentSN).Count > 1 Then
        crossFileDuplicates = crossFileDuplicates & "SN " & currentSN & " trouvé dans " & dictGlobalOverview(currentSN).Count & " fichiers :" & vbCrLf
        For Each logfile In dictGlobalOverview(currentSN).Keys
            crossFileDuplicates = crossFileDuplicates & "   - " & logfile & " (" & dictGlobalOverview(currentSN)(logfile) & " fois)" & vbCrLf
        Next
        crossFileDuplicates = crossFileDuplicates & vbCrLf
    End If
Next

results = results & String(80, "#") & vbCrLf
results = results & "SECTION ANALYSE DES DOUBLONS" & vbCrLf & String(80, "#") & vbCrLf & vbCrLf
results = results & "1. DOUBLONS DANS UN MÊME FICHIER :" & vbCrLf & String(40, "-") & vbCrLf
If sameFileDuplicates = "" Then results = results & "Aucun doublon interne trouvé." & vbCrLf Else results = results & sameFileDuplicates
results = results & vbCrLf
results = results & "2. DOUBLONS DANS DES FICHIERS DIFFÉRENTS :" & vbCrLf & String(40, "-") & vbCrLf
If crossFileDuplicates = "" Then results = results & "Aucun doublon inter-fichiers trouvé." & vbCrLf Else results = results & crossFileDuplicates

On Error Resume Next
Set fOut = objFSO.CreateTextFile(resultFile, True, False)
If Err.Number = 0 Then
    fOut.Write results
    fOut.Close
Else
    WScript.Echo "ERREUR : Impossible de créer le fichier d'inventaire OK."
End If
On Error GoTo 0

WScript.Echo "========================================" & vbCrLf & _
             "Inventaire des logs OK termine." & vbCrLf & _
             "Fichiers analyses      : " & fileCount & vbCrLf & _
             "Fichier genere         : Inventaire_Series_OK.txt" & vbCrLf & _
             "========================================"

Function SortArray(arr)
    Dim i, j, temp
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next
    Next
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
    CleanSerial = cleaned
End Function

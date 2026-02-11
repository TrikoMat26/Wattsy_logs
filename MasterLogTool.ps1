<#
    MasterLogTool.ps1
    Outil regroupant les analyses de logs ELISA (Selha).
    
    RESPECT DES CONSIGNES TECHNIQUES (LLM_Instructions.md) :
    1. Encodage Lecture/Ecriture : ANSI (Windows-1252).
    2. Nettoyage : Suppression des `0 (Null-bytes) et Trimming strict.
    3. Logique : Blocs délimités par "Datamatrix".
    4. Statuts : [PROD_OK], [PROD_ERROR], ou Incomplet.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION ET UTILITAIRES ---

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Définition stricte de l'encodage ANSI (Windows-1252) pour éviter les "caractères chinois"
$EncANSI = [System.Text.Encoding]::GetEncoding(1252)
$GapThreshold = 5   # Seuil d'écart pour la segmentation par lots/OF

# Fonction de nettoyage (Equivalent CleanSerial VBS)
Function Clean-Serial($val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return "" }
    # Suppression null bytes, =, tab, cr, lf, espaces
    $val = $val -replace "`0", "" `
        -replace "=", "" `
        -replace "`t", "" `
        -replace "`r", "" `
        -replace "`n", "" `
        -replace "\s+", ""
    return $val.Trim()
}

# Fonction unifiée d'extraction du SN (Supporte Ancien format # et Nouveau format :)
Function Extract-SN($line) {
    # Ancien format: ... Datamatrix: #2025#13104 ...
    if ($line -match "Datamatrix[:\s]*#.*?#(\d+)") {
        return Clean-Serial $matches[1]
    }
    # Nouveau format: ... Datamatrix: 043355 === (chiffres uniquement, pas d'URL)
    elseif ($line -match "Datamatrix[:\s]*(\d+)") {
        return Clean-Serial $matches[1]
    }
    return ""
}

# Fonction de Sanitization pour l'écriture (Garde seulement les chars ANSI < 255)
Function Sanitize-Text($str) {
    if ([string]::IsNullOrEmpty($str)) { return "" }
    # Remplace tout caractère dont le code est > 255 par un espace
    return [System.Text.RegularExpressions.Regex]::Replace($str, "[^\x00-\xFF]", " ")
}

# Fonction pour obtenir les fichiers logs
Function Get-LogFiles {
    return Get-ChildItem -Path $ScriptPath -Filter "ELISA_Prod_Log*.txt" | Sort-Object Name
}

# Fonction générique d'écriture de fichier (UTF-8 sans BOM pour les accents)
Function Write-ReportFile($path, $content) {
    try {
        # UTF-8 sans BOM préserve les accents et est compatible avec Notepad/VS Code
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
        $sw.Write($content)
        $sw.Close()
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'écriture du fichier : $path`n" + $_.Exception.Message, "Erreur", "OK", "Error")
        return $false
    }
}

# --- LOGIQUE MÉTIER (LES 4 SCRIPTS RECODÉS) ---

# 1. RECHERCHE SERIE KO
Function Run-RechercheSerie {
    $targetFile = Join-Path $ScriptPath "NumSerieKO.txt"
    if (-not (Test-Path $targetFile)) {
        return "ERREUR : Le fichier NumSerieKO.txt est introuvable."
    }

    $targets = @{}
    Get-Content $targetFile | ForEach-Object { $t = Clean-Serial $_; if ($t) { $targets[$t] = $true } }
    
    $results = @{} # Stockage en mémoire : Key=SN, Value=StringContent
    $counts = @{}  # Compteur d'occurrences
    $faults = @{}  # Key=SN, Value=liste ordonnée des défauts (un par bloc/occurrence)
    
    $logs = Get-LogFiles
    $totalLogs = $logs.Count
    
    # Scriptblock pour committer le défaut du bloc courant
    $CommitFault = {
        if ($currentSN -and $targets.ContainsKey($currentSN)) {
            if (-not $faults.ContainsKey($currentSN)) { $faults[$currentSN] = @() }
            if ($currentBlockFault) {
                $faults[$currentSN] += $currentBlockFault
            }
            else {
                $faults[$currentSN] += "Incomplet"
            }
        }
    }

    foreach ($log in $logs) {
        $lines = [System.IO.File]::ReadLines($log.FullName, $EncANSI)
        $currentSN = ""
        $currentBlockFault = ""
        
        foreach ($line in $lines) {
            $cleanLine = $line -replace "`0", ""
            
            if ($cleanLine.IndexOf("Datamatrix", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                # Commit le défaut du bloc précédent
                & $CommitFault
                $currentBlockFault = ""

                $sn = Extract-SN $cleanLine
                if ($sn) {
                    if ($targets.ContainsKey($sn)) {
                        $currentSN = $sn
                        if (-not $counts.ContainsKey($sn)) { $counts[$sn] = 0; $results[$sn] = [System.Text.StringBuilder]::new() }
                        $counts[$sn]++
                        
                        # En-tête obligatoire (LLM rules)
                        $sb = $results[$sn]
                        if ($sb.Length -gt 0) { [void]$sb.AppendLine() }
                        [void]$sb.AppendLine("-" * 40)
                        [void]$sb.AppendLine("Occurrence #$($counts[$sn]) (Fichier: $($log.Name))")
                        [void]$sb.AppendLine("-" * 40)
                        [void]$sb.AppendLine($cleanLine)
                    }
                    else {
                        $currentSN = ""
                    }
                }
                else { $currentSN = "" }
            }
            elseif ($currentSN) {
                [void]$results[$currentSN].AppendLine($cleanLine)
                # Détection du défaut : extraire les mots entre "[PROD_ERROR]: " et "fail"
                if (-not $currentBlockFault -and $cleanLine -match '\[PROD_ERROR\]:\s*(.+?)\s+fail\s*$') {
                    $currentBlockFault = $matches[1]
                }
                elseif (-not $currentBlockFault -and $cleanLine -match '\[PROD_ERROR\]:\s*(.+)$') {
                    # Cas où "fail" n'est pas présent : prendre tout après le tag
                    $currentBlockFault = $matches[1].Trim()
                }
            }
        }
        # Commit le défaut du dernier bloc du fichier
        & $CommitFault
        $currentBlockFault = ""
        $currentSN = ""
    }

    # Ecriture avec noms enrichis des défauts
    $cnt = 0
    foreach ($sn in $results.Keys) {
        # Construction du suffixe : _défaut1_défaut2_...
        $faultSuffix = ""
        if ($faults.ContainsKey($sn) -and $faults[$sn].Count -gt 0) {
            $faultSuffix = "_" + ($faults[$sn] -join "_")
        }
        $fPath = Join-Path $ScriptPath "$sn$faultSuffix.txt"
        $content = Sanitize-Text $results[$sn].ToString()
        Write-ReportFile $fPath $content
        $cnt++
    }
    # Bilan : trouvés vs non trouvés
    $notFound = @()
    foreach ($t in $targets.Keys) {
        if (-not $results.ContainsKey($t)) { $notFound += $t }
    }
    $notFound = $notFound | Sort-Object
    $msg = "Termine. $cnt fichier(s) genere(s) sur $($targets.Count) SN recherche(s)."
    if ($notFound.Count -gt 0) {
        $msg += "`n$($notFound.Count) SN non trouve(s) dans les logs : " + ($notFound -join ", ")
    }

    # Ajout du bilan à la fin de Inventaire_Series_OK.txt
    $okPath = Join-Path $ScriptPath "Inventaire_Series_OK.txt"
    $bilanSb = [System.Text.StringBuilder]::new()
    [void]$bilanSb.AppendLine()
    [void]$bilanSb.AppendLine("#" * 80)
    [void]$bilanSb.AppendLine("BILAN RECHERCHE SERIE KO (NumSerieKO.txt)")
    [void]$bilanSb.AppendLine("Genere le : $(Get-Date)")
    [void]$bilanSb.AppendLine("#" * 80)
    [void]$bilanSb.AppendLine()
    [void]$bilanSb.AppendLine("SN recherches : $($targets.Count)")
    [void]$bilanSb.AppendLine("SN trouves (fichiers generes) : $cnt")
    if ($notFound.Count -gt 0) {
        [void]$bilanSb.AppendLine("SN non trouves ($($notFound.Count)) : " + ($notFound -join ", "))
    }
    else {
        [void]$bilanSb.AppendLine("SN non trouves : 0 (tous trouves)")
    }
    [void]$bilanSb.AppendLine()

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    if (Test-Path -LiteralPath $okPath) {
        [System.IO.File]::AppendAllText($okPath, $bilanSb.ToString(), $utf8NoBom)
    }
    else {
        [System.IO.File]::WriteAllText($okPath, $bilanSb.ToString(), $utf8NoBom)
    }

    return $msg
}


# 2. INVENTAIRE SERIES
Function Run-InventaireSeries {
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("INVENTAIRE DES NUMEROS DE SERIE PAR FICHIER LOG")
    [void]$sb.AppendLine("Genere le : $(Get-Date)")
    [void]$sb.AppendLine("=" * 60)
    [void]$sb.AppendLine()

    $logs = Get-LogFiles
    $totalLogsLogs = $logs.Count

    foreach ($log in $logs) {
        $lines = [System.IO.File]::ReadLines($log.FullName, $EncANSI)
        $dictInFile = [System.Collections.Generic.HashSet[string]]::new()
        
        foreach ($line in $lines) {
            $cleanLine = $line -replace "`0", ""
            if ($cleanLine.IndexOf("Datamatrix", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $sn = Extract-SN $cleanLine
                if ($sn) { [void]$dictInFile.Add($sn) }
            }
        }

        [void]$sb.AppendLine("FICHIER : $($log.Name)")
        [void]$sb.AppendLine("Nombre de SN uniques : $($dictInFile.Count)")
        [void]$sb.AppendLine("Liste : " + ($dictInFile -join ", "))
        [void]$sb.AppendLine("-" * 60)
        [void]$sb.AppendLine()
    }

    $outPath = Join-Path $ScriptPath "Inventaire_Series.txt"
    Write-ReportFile $outPath (Sanitize-Text $sb.ToString())
    return "Terminé. Fichier généré : Inventaire_Series.txt"
}

# 3. INVENTAIRE SERIES OK + DOUBLONS
Function Run-InventaireSeriesOK {
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("INVENTAIRE DES NUMEROS DE SERIE AVEC TEST OK (SUCCESSFUL)")
    [void]$sb.AppendLine("Critere : Contient [PROD_OK]")
    [void]$sb.AppendLine("=" * 80)
    [void]$sb.AppendLine()

    $logs = Get-LogFiles
    
    # Doublons Cross-files
    $globalOverview = @{} # Key=SN, Val=Dict(File->Count)
    $sameFileDupText = [System.Text.StringBuilder]::new()

    foreach ($log in $logs) {
        $lines = [System.IO.File]::ReadLines($log.FullName, $EncANSI)
        
        $snCountsInFile = [System.Collections.Specialized.OrderedDictionary]::new()
        
        $currentSN = ""
        $isBlockOK = $false

        foreach ($line in $lines) {
            $cleanLine = $line -replace "`0", ""
            
            if ($cleanLine.IndexOf("Datamatrix", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                # Fin bloc précédent
                if ($currentSN -ne "" -and $isBlockOK) {
                    if (-not $snCountsInFile.Contains($currentSN)) { $snCountsInFile[$currentSN] = 0 }
                    $snCountsInFile[$currentSN]++

                    if (-not $globalOverview.Contains($currentSN)) { $globalOverview[$currentSN] = @{} }
                    if (-not $globalOverview[$currentSN].Contains($log.Name)) { $globalOverview[$currentSN][$log.Name] = 0 }
                    $globalOverview[$currentSN][$log.Name]++
                }

                # Nouveau bloc
                $sn = Extract-SN $cleanLine
                if ($sn) {
                    $currentSN = $sn
                    $isBlockOK = ($cleanLine.IndexOf("[PROD_OK]", [StringComparison]::OrdinalIgnoreCase) -ge 0)
                }
                else { $currentSN = "" }

            }
            elseif ($currentSN) {
                if (-not $isBlockOK -and $cleanLine.IndexOf("[PROD_OK]", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $isBlockOK = $true
                }
            }
        }
        # Dernier bloc
        if ($currentSN -ne "" -and $isBlockOK) {
            if (-not $snCountsInFile.Contains($currentSN)) { $snCountsInFile[$currentSN] = 0 }
            $snCountsInFile[$currentSN]++

            if (-not $globalOverview.Contains($currentSN)) { $globalOverview[$currentSN] = @{} }
            if (-not $globalOverview[$currentSN].Contains($log.Name)) { $globalOverview[$currentSN][$log.Name] = 0 }
            $globalOverview[$currentSN][$log.Name]++
        }

        # Rapport par fichier
        if ($snCountsInFile.Count -gt 0) {
            [void]$sb.AppendLine("FICHIER : $($log.Name)")
            [void]$sb.AppendLine("Nombre de SN OK (distincts) : $($snCountsInFile.Count)")
            
            # Ordre d'apparition
            [void]$sb.Append(" - Par ordre d'apparition : ")
            [void]$sb.AppendLine(($snCountsInFile.Keys -join ", "))
            
            # Ordre croissant
            $sortedKeys = $snCountsInFile.Keys | Sort-Object
            [void]$sb.Append(" - Par ordre croissant   : ")
            [void]$sb.AppendLine(($sortedKeys -join ", "))
            
            # Doublons internes
            $localDups = [System.Text.StringBuilder]::new()
            foreach ($k in $snCountsInFile.Keys) {
                if ($snCountsInFile[$k] -gt 1) {
                    [void]$localDups.AppendLine("   - $k ($($snCountsInFile[$k]) fois)")
                }
            }
            if ($localDups.Length -gt 0) {
                [void]$sameFileDupText.AppendLine("Dans $($log.Name) :")
                [void]$sameFileDupText.Append($localDups.ToString())
                [void]$sameFileDupText.AppendLine()
            }
            [void]$sb.AppendLine("-" * 80)
            [void]$sb.AppendLine()
        }
    }

    # Section Doublons
    [void]$sb.AppendLine("#" * 80)
    [void]$sb.AppendLine("SECTION ANALYSE DES DOUBLONS")
    [void]$sb.AppendLine("#" * 80)
    [void]$sb.AppendLine()
    
    [void]$sb.AppendLine("1. DOUBLONS DANS UN MÊME FICHIER :")
    [void]$sb.AppendLine("-" * 40)
    if ($sameFileDupText.Length -eq 0) { [void]$sb.AppendLine("Aucun.") } else { [void]$sb.Append($sameFileDupText.ToString()) }
    
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("2. DOUBLONS DANS DES FICHIERS DIFFÉRENTS :")
    [void]$sb.AppendLine("-" * 40)
    
    $hasCross = $false
    foreach ($sn in $globalOverview.Keys) {
        if ($globalOverview[$sn].Count -gt 1) {
            $hasCross = $true
            [void]$sb.AppendLine("SN $sn trouvé dans $($globalOverview[$sn].Count) fichiers :")
            foreach ($f in $globalOverview[$sn].Keys) {
                [void]$sb.AppendLine("   - $f ($($globalOverview[$sn][$f]) fois)")
            }
            [void]$sb.AppendLine()
        }
    }
    if (-not $hasCross) { [void]$sb.AppendLine("Aucun.") }

    # Section Liste Globale Confondus
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("#" * 80)
    [void]$sb.AppendLine("LISTE GLOBALE DES SN OK (CONFONDUS - TOUS FICHIERS)")
    [void]$sb.AppendLine("#" * 80)
    [void]$sb.AppendLine()
    
    $allUniqueSN = $globalOverview.Keys | Sort-Object
    [void]$sb.AppendLine("Nombre total de SN uniques OK : $($allUniqueSN.Count)")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("Liste triée par ordre croissant :")
    [void]$sb.AppendLine(($allUniqueSN -join ", "))
    [void]$sb.AppendLine()

    # --- SEGMENTATION PAR LOTS / OF (issue de Liste_OF.ps1) ---
    [void]$sb.AppendLine("#" * 80)
    [void]$sb.AppendLine("ANALYSE PAR SEGMENTS (LOTS / OF)")
    [void]$sb.AppendLine("Seuil d'ecart : $GapThreshold (nouveau segment si ecart > $GapThreshold)")
    [void]$sb.AppendLine("#" * 80)
    [void]$sb.AppendLine()

    if ($allUniqueSN.Count -gt 0) {
        # Construction widthMap : SN(int) -> largeur texte max observée
        $widthMap = @{}
        foreach ($s in $allUniqueSN) {
            try { $n = [int]$s } catch { continue }
            $len = $s.Length
            if (-not $widthMap.ContainsKey($n)) {
                $widthMap[$n] = $len
            }
            elseif ($len -gt [int]$widthMap[$n]) {
                $widthMap[$n] = $len
            }
        }

        # Tri numérique
        $nums = $widthMap.Keys | Sort-Object { [int]$_ }

        if ($nums -and $nums.Count -gt 0) {
            # Segmentation
            $segments = @()
            $currentSeg = @($nums[0])

            for ($i = 1; $i -lt $nums.Count; $i++) {
                $prev = [int]$nums[$i - 1]
                $cur = [int]$nums[$i]
                $gap = $cur - $prev

                if ($gap -gt $GapThreshold) {
                    $segments += , $currentSeg
                    $currentSeg = @($cur)
                }
                else {
                    $currentSeg += $cur
                }
            }
            $segments += , $currentSeg

            # Rendu par segment + collecte globale des manquants
            $segId = 0
            $allMissingSN = @()
            foreach ($seg in $segments) {
                $segId++

                $segStart = [int]$seg[0]
                $segEnd = [int]$seg[$seg.Count - 1]
                $presentCount = $seg.Count

                # Largeur d'affichage du segment = max des largeurs observées
                $segWidth = 0
                foreach ($n in $seg) {
                    $w = [int]$widthMap[[int]$n]
                    if ($w -gt $segWidth) { $segWidth = $w }
                }
                if ($segWidth -lt 1) { $segWidth = 6 }

                # Hashtable de présence
                $present = @{}
                foreach ($n in $seg) { $present[[int]$n] = $true }

                # Calcul des manquants
                $missing = @()
                for ($n = $segStart; $n -le $segEnd; $n++) {
                    if (-not $present.ContainsKey($n)) {
                        $missing += $n.ToString("D$segWidth")
                        $allMissingSN += $n.ToString("D$segWidth")
                    }
                }

                $rangeStr = "{0} - {1}" -f $segStart.ToString("D$segWidth"), $segEnd.ToString("D$segWidth")

                if ($missing.Count -gt 0) {
                    [void]$sb.AppendLine(("segment {0} : {1}, present={2}, missing={3} ({4})" -f $segId, $rangeStr, $presentCount, $missing.Count, ($missing -join ", ")))
                }
                else {
                    [void]$sb.AppendLine(("segment {0} : {1}, present={2}, missing=0" -f $segId, $rangeStr, $presentCount))
                }
            }

            [void]$sb.AppendLine()
            [void]$sb.AppendLine("Total : $($segments.Count) segment(s) detecte(s).")
        }
        else {
            [void]$sb.AppendLine("Aucun numero exploitable pour la segmentation.")
        }
    }
    else {
        [void]$sb.AppendLine("Aucun SN OK trouve, segmentation impossible.")
    }
    [void]$sb.AppendLine()

    $outPath = Join-Path $ScriptPath "Inventaire_Series_OK.txt"
    Write-ReportFile $outPath (Sanitize-Text $sb.ToString())

    # --- Proposition d'export des manquants vers NumSerieKO.txt ---
    if ($allMissingSN -and $allMissingSN.Count -gt 0) {
        $msgText = "$($allMissingSN.Count) numero(s) manquant(s) detecte(s) dans les segments :`n`n" + ($allMissingSN -join ", ") + "`n`nVoulez-vous les ajouter dans NumSerieKO.txt ?"
        $answer = [System.Windows.Forms.MessageBox]::Show($msgText, "Numeros manquants", "YesNo", "Question")

        if ($answer -eq "Yes") {
            $koPath = Join-Path $ScriptPath "NumSerieKO.txt"
            # Lire le contenu existant (ou vide si fichier inexistant)
            $existingContent = ""
            if (Test-Path -LiteralPath $koPath) {
                $existingContent = (Get-Content -LiteralPath $koPath -Raw -ErrorAction SilentlyContinue)
                if (-not $existingContent) { $existingContent = "" }
            }
            # Ajouter les manquants (un par ligne) à la suite
            $newEntries = $allMissingSN -join "`r`n"
            if ($existingContent.Length -gt 0 -and -not $existingContent.EndsWith("`n")) {
                $newEntries = "`r`n" + $newEntries
            }
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::AppendAllText($koPath, $newEntries + "`r`n", $utf8NoBom)
            return "Termine. Fichier genere : Inventaire_Series_OK.txt`n$($allMissingSN.Count) manquant(s) ajoute(s) dans NumSerieKO.txt."
        }
    }

    return "Termine. Fichier genere : Inventaire_Series_OK.txt"
}

# 4. HISTORIQUE TESTS
Function Run-HistoriqueTests {
    $dictHist = [System.Collections.Generic.SortedDictionary[string, string]]::new()
    
    $logs = Get-LogFiles
    
    foreach ($log in $logs) {
        $lines = [System.IO.File]::ReadLines($log.FullName, $EncANSI)
        
        $currentSN = ""
        $currentStatus = "Test Incomplet"
        
        # Procédure interne pour éviter duplication code
        $CommitBlock = {
            if ($currentSN) {
                $info = "   -> " + ($currentStatus.PadRight(50)).Substring(0, 50) + " | Fichier: $($log.Name)`r`n"
                if (-not $dictHist.ContainsKey($currentSN)) { $dictHist[$currentSN] = "" }
                $dictHist[$currentSN] += $info
            }
        }

        foreach ($line in $lines) {
            $cleanLine = $line -replace "`0", ""
            
            if ($cleanLine.IndexOf("Datamatrix", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                # Commit previous
                & $CommitBlock
                
                # Start new
                $sn = Extract-SN $cleanLine
                if ($sn) {
                    $currentSN = $sn
                    $currentStatus = "Test Incomplet"
                    # Check instantané
                    if ($cleanLine.IndexOf("[PROD_OK]", [StringComparison]::OrdinalIgnoreCase) -ge 0) { $currentStatus = "Test OK" }
                    elseif ($cleanLine.IndexOf("[PROD_ERROR]", [StringComparison]::OrdinalIgnoreCase) -ge 0) { $currentStatus = $cleanLine.Trim() }
                }
                else { $currentSN = "" }
            }
            elseif ($currentSN) {
                if ($cleanLine.IndexOf("[PROD_OK]", [StringComparison]::OrdinalIgnoreCase) -ge 0) { 
                    $currentStatus = "Test OK" 
                }
                elseif ($cleanLine.IndexOf("[PROD_ERROR]", [StringComparison]::OrdinalIgnoreCase) -ge 0) { 
                    $currentStatus = $cleanLine.Trim() 
                }
            }
        }
        # Commit last
        & $CommitBlock
    }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("HISTORIQUE COMPLET DES PASSAGES")
    [void]$sb.AppendLine("Genere le : $(Get-Date)")
    [void]$sb.AppendLine("Fichiers analyses : $($logs.Count)")
    [void]$sb.AppendLine("=" * 80)
    [void]$sb.AppendLine()

    foreach ($sn in $dictHist.Keys) {
        [void]$sb.AppendLine("SN: $sn")
        [void]$sb.Append($dictHist[$sn]) 
        [void]$sb.AppendLine("-" * 60)
    }

    $outPath = Join-Path $ScriptPath "Historique_Tests_Complet.txt"
    Write-ReportFile $outPath (Sanitize-Text $sb.ToString())
    return "Terminé. Fichier généré : Historique_Tests_Complet.txt"
}


# --- INTERFACE GRAPHIQUE (WINFORMS) ---

$Form = New-Object Windows.Forms.Form
$Form.Text = "Selha - ELISA Log Analyzer Tool"
$Form.Size = New-Object Drawing.Size(800, 600)
$Form.StartPosition = "CenterScreen"
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Layout principal
$Table = New-Object Windows.Forms.TableLayoutPanel
$Table.Dock = "Fill"
$Table.ColumnCount = 2
$Table.RowCount = 2
$Table.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 40)))
$Table.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 60)))
$Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 60)))
$Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 40)))
$Form.Controls.Add($Table)

# Panel Gauche (Sélection Scripts)
$PanelLeft = New-Object Windows.Forms.FlowLayoutPanel
$PanelLeft.Dock = "Fill"
$PanelLeft.FlowDirection = "TopDown"
$PanelLeft.WrapContents = $false 
$PanelLeft.Padding = New-Object Windows.Forms.Padding(10)
$Table.Controls.Add($PanelLeft, 0, 0)

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = "Choisir une action :"
$lblTitle.AutoSize = $true
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$PanelLeft.Controls.Add($lblTitle)

$rb1 = New-Object Windows.Forms.RadioButton
$rb1.Text = "1. Inventaire Validés (OK & Doublons)"
$rb1.AutoSize = $true
$rb1.Checked = $true
$rb1.Tag = "Desc: Liste les SN OK ([PROD_OK]), analyse les doublons, et effectue une segmentation par lots (Analyse OF) avec détection des manquants."
$PanelLeft.Controls.Add($rb1)

$rb2 = New-Object Windows.Forms.RadioButton
$rb2.Text = "2. Extraction par Liste (NumSerieKO)"
$rb2.AutoSize = $true
$rb2.Tag = "Desc: Extrait les logs complets pour chaque numéro de série présent dans 'NumSerieKO.txt'. Crée un fichier par SN avec le défaut dans le nom."
$PanelLeft.Controls.Add($rb2)

$rb3 = New-Object Windows.Forms.RadioButton
$rb3.Text = "3. Inventaire Global (Tous)"
$rb3.AutoSize = $true
$rb3.Tag = "Desc: Liste tous les numéros de série trouvés dans chaque fichier log, sans filtrage."
$PanelLeft.Controls.Add($rb3)

$rb4 = New-Object Windows.Forms.RadioButton
$rb4.Text = "4. Historique Complet (Traçabilité)"
$rb4.AutoSize = $true
$rb4.Tag = "Desc: Trace chronologiquement tous les passages de chaque SN. Indique si OK, Erreur précise ou Incomplet."
$PanelLeft.Controls.Add($rb4)

$descBox = New-Object Windows.Forms.Label
$descBox.Text = $rb1.Tag
$descBox.AutoSize = $false
$descBox.Width = 280
$descBox.Height = 100
$descBox.BorderStyle = "Fixed3D"
$descBox.Padding = New-Object Windows.Forms.Padding(5)
$descBox.Margin = New-Object Windows.Forms.Padding(0, 20, 0, 0)
$PanelLeft.Controls.Add($descBox)

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = " EXÉCUTER LE SCRIPT "
$btnRun.Height = 50
$btnRun.Width = 280
$btnRun.BackColor = [System.Drawing.Color]::LightGreen
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$PanelLeft.Controls.Add($btnRun)

# Panel Droit (Résultats et Fichiers)
$PanelRight = New-Object Windows.Forms.Panel
$PanelRight.Dock = "Fill"
$PanelRight.Padding = New-Object Windows.Forms.Padding(10)
$Table.Controls.Add($PanelRight, 1, 0)
$Table.SetRowSpan($PanelRight, 2) # Prend toute la hauteur à droite

$lblFiles = New-Object Windows.Forms.Label
$lblFiles.Text = "Fichiers Générés (.txt) :"
$lblFiles.AutoSize = $true
$lblFiles.Top = 10
$PanelRight.Controls.Add($lblFiles)

$lstFiles = New-Object Windows.Forms.ListBox
$lstFiles.Top = 35
$lstFiles.Left = 0
$lstFiles.Width = 450
$lstFiles.Height = 400
$lstFiles.Anchor = "Top, Left, Right"
$PanelRight.Controls.Add($lstFiles)

$btnOpen = New-Object Windows.Forms.Button
$btnOpen.Text = "Ouvrir avec Notepad"
$btnOpen.Top = 445
$btnOpen.Width = 150
$PanelRight.Controls.Add($btnOpen)

$btnRefresh = New-Object Windows.Forms.Button
$btnRefresh.Text = "Actualiser la liste"
$btnRefresh.Top = 445
$btnRefresh.Left = 160
$btnRefresh.Width = 150
$PanelRight.Controls.Add($btnRefresh)

# Panel Bas-Gauche (Console Log)
$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Dock = "Fill"
$txtLog.ReadOnly = $true
$txtLog.BackColor = "Black"
$txtLog.ForeColor = "White"
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$Table.Controls.Add($txtLog, 0, 1)

# --- EVENTS ---

# Mise à jour description
$rb1.Add_CheckedChanged({ if ($rb1.Checked) { $descBox.Text = $rb1.Tag } })
$rb2.Add_CheckedChanged({ if ($rb2.Checked) { $descBox.Text = $rb2.Tag } })
$rb3.Add_CheckedChanged({ if ($rb3.Checked) { $descBox.Text = $rb3.Tag } })
$rb4.Add_CheckedChanged({ if ($rb4.Checked) { $descBox.Text = $rb4.Tag } })

# Fonction Refresh liste fichiers
$RefreshFiles = {
    $lstFiles.Items.Clear()
    $files = Get-ChildItem -Path $ScriptPath -Filter "*.txt" | Where-Object { $_.Name -notlike "ELISA_Prod_Log*" -and $_.Name -ne "NumSerieKO.txt" -and $_.Name -ne "InfoScript.txt" } | Sort-Object LastWriteTime -Descending
    foreach ($f in $files) { [void]$lstFiles.Items.Add($f.Name) }
}
$btnRefresh.Add_Click($RefreshFiles)

# Initial Load
& $RefreshFiles

# Fonction Log Console
Function Log-Console($msg) {
    $txtLog.AppendText("[$([DateTime]::Now.ToString('HH:mm:ss'))] $msg`r`n")
}

# Exécution Script
$btnRun.Add_Click({
        $btnRun.Enabled = $false
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
        Log-Console "Démarrage du traitement..."
        [System.Windows.Forms.Application]::DoEvents()

        $res = ""
        try {
            if ($rb1.Checked) { Log-Console "Script: Inventaire OK..."; $res = Run-InventaireSeriesOK }
            elseif ($rb2.Checked) { Log-Console "Script: Recherche KO..."; $res = Run-RechercheSerie }
            elseif ($rb3.Checked) { Log-Console "Script: Inventaire..."; $res = Run-InventaireSeries }
            elseif ($rb4.Checked) { Log-Console "Script: Historique..."; $res = Run-HistoriqueTests }
        
            Log-Console $res
        }
        catch {
            Log-Console "ERREUR FATALE : $_"
        }

        & $RefreshFiles
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRun.Enabled = $true
    })

# Ouvrir Fichier
$btnOpen.Add_Click({
        if ($lstFiles.SelectedItem) {
            $f = Join-Path $ScriptPath $lstFiles.SelectedItem
            Start-Process "notepad.exe" $f
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner un fichier dans la liste.", "Info")
        }
    })

# Lancement
Log-Console "Prêt. Répertoire : $ScriptPath"
Log-Console "Encodage actif : ANSI (1252)"

$Form.Add_Shown({ $Form.Activate() })
[void]$Form.ShowDialog()

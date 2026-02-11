<#
PS 5.1 - Script robuste

Input :  Liste_OF.txt (même dossier que ce script)
Output:  Liste_OF_traité.txt (même dossier)

Règle segmentation : nouveau segment si (écart > 5) entre deux n° consécutifs triés.
#>

# -----------------------------
# Paramètres
# -----------------------------
$GapThreshold = 5
$Exclude = @()   # ex: @("043246")

# -----------------------------
# Chemins
# -----------------------------
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$InputFile  = Join-Path $ScriptDir "Liste_OF.txt"
$OutputFile = Join-Path $ScriptDir "Liste_OF_traité.txt"

if (-not (Test-Path -LiteralPath $InputFile)) {
    throw "Fichier introuvable : $InputFile"
}

# -----------------------------
# Lecture + extraction nombres
# -----------------------------
$text = Get-Content -LiteralPath $InputFile -Raw -ErrorAction Stop

# Récupère tous les groupes de chiffres
$matches = [regex]::Matches($text, '\d+')

# Map: serial(int) -> largeur max observée (pour garder 043355 vs 13088)
$widthMap = @{}

foreach ($m in $matches) {
    $s = $m.Value
    try { $n = [int]$s } catch { continue }
    if ($Exclude -contains $s) { continue }   # exclusion par forme texte
    $len = $s.Length
    if (-not $widthMap.ContainsKey($n)) {
        $widthMap[$n] = $len
    } elseif ($len -gt [int]$widthMap[$n]) {
        $widthMap[$n] = $len
    }
}

$nums = $widthMap.Keys | Sort-Object { [int]$_ }

if (-not $nums -or $nums.Count -eq 0) {
    "Aucun numéro de série détecté dans $InputFile" | Set-Content -LiteralPath $OutputFile -Encoding UTF8
    Write-Output "Aucun numéro de série détecté dans $InputFile"
    Write-Output "Résultat enregistré dans : $OutputFile"
    exit
}

# -----------------------------
# Segmentation
# -----------------------------
# segments = tableau de tableaux (int[])
$segments = @()
$current = @($nums[0])

for ($i = 1; $i -lt $nums.Count; $i++) {
    $prev = [int]$nums[$i - 1]
    $cur  = [int]$nums[$i]
    $gap  = $cur - $prev

    if ($gap -gt $GapThreshold) {
        $segments += ,$current
        $current = @($cur)
    } else {
        $current += $cur
    }
}
$segments += ,$current

function Fmt([int]$n, [int]$w) { $n.ToString("D$w") }

# -----------------------------
# Analyse + rendu
# -----------------------------
$outLines = New-Object System.Collections.Generic.List[string]
$segId = 0

foreach ($seg in $segments) {
    $segId++

    $start = [int]$seg[0]
    $end   = [int]$seg[$seg.Count - 1]
    $presentCount = $seg.Count

    # largeur d'affichage du segment = max des largeurs observées sur ses valeurs
    $segWidth = 0
    foreach ($n in $seg) {
        $w = [int]$widthMap[[int]$n]
        if ($w -gt $segWidth) { $segWidth = $w }
    }
    if ($segWidth -lt 1) { $segWidth = 6 }

    # présence
    $present = @{}
    foreach ($n in $seg) { $present[[int]$n] = $true }

    # manquants
    $missing = @()
    for ($n = $start; $n -le $end; $n++) {
        if (-not $present.ContainsKey($n)) {
            $missing += (Fmt $n $segWidth)
        }
    }

    $range = "{0}–{1}" -f (Fmt $start $segWidth), (Fmt $end $segWidth)

    if ($missing.Count -gt 0) {
        $outLines.Add(("segment {0} : {1}, present={2}, missing={3} ({4})" -f $segId, $range, $presentCount, $missing.Count, ($missing -join ", "))) | Out-Null
    } else {
        $outLines.Add(("segment {0} : {1}, present={2}, missing=0" -f $segId, $range, $presentCount)) | Out-Null
    }
}

# -----------------------------
# Sauvegarde + affichage
# -----------------------------
$outLines.ToArray() | Set-Content -LiteralPath $OutputFile -Encoding UTF8
$outLines.ToArray()
"Résultat enregistré dans : $OutputFile"

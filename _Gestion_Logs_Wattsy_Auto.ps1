
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ======================================================
# Get-Wattsy-Logs.ps1 (MENU)
# - WinSCP portable (sans droits admin)
# - Menu : 1) Horodatage  2) Scan  3) Téléchargement  4) Archivage  0) Quitter
# - Téléchargement: nouveaux + modifiés + contrôle contenu (MD5)
# - Pas d’erreur si aucun fichier (tout archivé)
# - Logs d’exécution horodatés
# ======================================================

# ----------------------
# PARAMÈTRES
# ----------------------
$WinSCP = "C:\Users\KKAYZ\Tools\WinSCP\WinSCP.com"

$HostName = "piwio.local"
$Port = 22
$UserName = "piwio"
$Password = "piwio"

$RemoteLogsDir = "/home/piwio/banc_test_lte/logs"
$FileMask = "ELISA_Prod_Log_*"

$LocalDir = "C:\Users\KKAYZ\OneDrive - SELHA GROUP\Clients\wattsy\logs"

# HostKeyFingerprint (format WinSCP)
$HostKeyFingerprint = "ssh-ed25519 255 6smJTCgUEDSpmIGlpojJB6Ppuajmv9MEGqAxKvxpjJU"

# Robustesse WinSCP
$RawSettings = "SendBuf=0"
$OpenTimeout = 300

# Option horodatage
$TimeToleranceSeconds = 5

# ----------------------
# LOGS
# ----------------------
$RunStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ExecLogDir = Join-Path $LocalDir "_logs_exec"
$PsLog = Join-Path $ExecLogDir ("Get-Wattsy-Logs_" + $RunStamp + ".log")
$WinSCPLog = Join-Path $ExecLogDir ("WinSCP_" + $RunStamp + ".log")

# ----------------------
# VÉRIFICATIONS
# ----------------------
if (-not (Test-Path -LiteralPath $WinSCP)) { throw "WinSCP.com introuvable : $WinSCP" }
if (-not (Test-Path -LiteralPath $LocalDir)) { New-Item -ItemType Directory -Path $LocalDir | Out-Null }
if (-not (Test-Path -LiteralPath $ExecLogDir)) { New-Item -ItemType Directory -Path $ExecLogDir | Out-Null }

Start-Transcript -Path $PsLog -Append | Out-Null

# ======================================================
# OUTILS GENERIQUES
# ======================================================
function New-TempFile([string]$ext = ".txt") {
    [System.IO.Path]::Combine($env:TEMP, ([guid]::NewGuid().ToString("N") + $ext))
}

function Invoke-WinSCP([string]$scriptText) {
    $tmp = New-TempFile ".winscp.txt"
    Set-Content -LiteralPath $tmp -Value $scriptText -Encoding ASCII
    try {
        $out = & $WinSCP /ini=nul /log="$WinSCPLog" /loglevel=1 /script="$tmp" 2>&1 | Out-String
        $code = $LASTEXITCODE
        return [pscustomobject]@{ ExitCode = $code; Output = $out }
    }
    finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

function New-OpenLine {
    "open sftp://${UserName}:${Password}@${HostName}:${Port}/ -timeout=$OpenTimeout -hostkey=""$HostKeyFingerprint"" -rawsettings $RawSettings"
}

function Escape-ForSingleQuotes([string]$s) {
    # Pour sh -lc '...': remplace ' par '"'"'
    return ($s -replace "'", "'""'""'")
}

function Invoke-RemoteShell([string]$cmd) {
    $safe = Escape-ForSingleQuotes $cmd
    $script = @"
option batch abort
option confirm off
$(New-OpenLine)
call sh -lc '$safe'
exit
"@
    Invoke-WinSCP $script
}

function Invoke-RemoteShellMulti([string[]]$cmds) {
    $callLines = ($cmds | ForEach-Object {
            $safe = Escape-ForSingleQuotes $_
            "call sh -lc '$safe'"
        }) -join "`r`n"

    $script = @"
option batch abort
option confirm off
$(New-OpenLine)
$callLines
exit
"@
    Invoke-WinSCP $script
}

# ======================================================
# UI MENU
# ======================================================
function Show-Menu {
    Write-Host ""
    Write-Host "============================================"
    Write-Host "Wattsy Logs - MENU"
    Write-Host "--------------------------------------------"
    Write-Host "1. Vérification horodatage (PC vs Pi) + proposition mise à l'heure"
    Write-Host "2. Scan des fichiers logs (compte + répertoires + liste)"
    Write-Host "3. Téléchargement des fichiers logs"
    Write-Host "4. Archivage des fichiers logs sur le Pi"
    Write-Host "0. Quitter"
    Write-Host "============================================"
}

function Show-ActionHeader([string]$title) {
    Write-Host ""
    Write-Host "▶ $title"
    Write-Host "--------------------------------------------"
}

function Show-ActionFooter {
    Write-Host "--------------------------------------------"
    Write-Host "✔ Action terminée"
    [void](Read-Host "Appuyer sur Entrée pour revenir au menu")
}

# ======================================================
# 1) HORODATAGE  (MODIFIÉ RTC UNIQUEMENT)
# - Lecture RTC via /home/piwio/banc_test_lte/test_rtc_2.py
# - Mise à l’heure = écriture RTC (inspirée one_shot_rtc_set.py)
# - Après MAJ RTC, proposition de reboot (nécessaire pour que le système se cale sur le RTC au démarrage)
# ======================================================

$RemoteRtcDir = "/home/piwio/banc_test_lte"
$RemoteRtcTestPy = "$RemoteRtcDir/test_rtc_2.py"

function Get-RtcDateTime {
    $r = Invoke-RemoteShell "python3 '$RemoteRtcTestPy'"
    if ($r.ExitCode -ne 0) {
        throw "Erreur lecture RTC via test_rtc_2.py (code $($r.ExitCode)). Voir: $WinSCPLog`n$($r.Output)"
    }

    # Attendu: "RTC Time: 2026-02-03 10:34:20"
    if ($r.Output -match '(?m)^\s*RTC Time:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s*$') {
        return [datetime]::ParseExact($matches[1], "yyyy-MM-dd HH:mm:ss", $null)
    }

    throw "Impossible d'extraire la date RTC. Sortie:`n$($r.Output)"
}

function Get-PiUnixTime {
    # (Conservé comme API interne) => retourne le temps UNIX du RTC
    $rtc = Get-RtcDateTime
    return [int64]([DateTimeOffset]$rtc).ToUnixTimeSeconds()
}

function Get-PiHumanTime {
    # (Conservé comme API interne) => retourne une chaîne lisible du RTC
    $rtc = Get-RtcDateTime
    return ($rtc.ToString("yyyy-MM-dd HH:mm:ss") + " (RTC)")
}

function Set-RtcFromPcNow([datetime]$pcNow) {

    # DS3231 weekday: 1=Monday ... 7=Sunday
    $dow = [int]$pcNow.DayOfWeek
    $weekday = if ($dow -eq 0) { 7 } else { $dow }

    $s = $pcNow.Second
    $mi = $pcNow.Minute
    $h = $pcNow.Hour
    $d = $pcNow.Day
    $mo = $pcNow.Month
    $yy = $pcNow.Year - 2000

    # Code python (SANS apostrophes) -> on l'encapsule ensuite dans des quotes simples côté shell
    $py = "from smbus2 import SMBus; b=lambda v:((v//10)<<4)+(v%10); bus=SMBus(1); bus.write_i2c_block_data(0x68,0,[b($s),b($mi),b($h),b($weekday),b($d),b($mo),b($yy)]); print(""RTC time set."")"

    # IMPORTANT: python -c '...'(quotes simples côté shell)
    # Invoke-RemoteShell échappe déjà les quotes simples via Escape-ForSingleQuotes
    $cmd = "python3 -c '$py'"

    $r = Invoke-RemoteShell $cmd
    if ($r.ExitCode -ne 0) {
        throw "Échec écriture RTC (code $($r.ExitCode)). Sortie:`n$($r.Output)"
    }
}

function Check-Time-And-OfferSync {
    $pcNow = Get-Date
    $pcUnix = [int64]([DateTimeOffset]$pcNow).ToUnixTimeSeconds()
    $pcHuman = $pcNow.ToString("yyyy-MM-dd HH:mm:ss")
    $pcTz = (Get-TimeZone).Id

    # Pi = RTC (pas l’horloge système)
    $rtcHuman = Get-PiHumanTime
    $rtcUnix = Get-PiUnixTime

    $diff = [math]::Abs($pcUnix - $rtcUnix)

    Write-Host "PC : $pcHuman ($pcTz) | unix=$pcUnix"
    Write-Host "PI : $rtcHuman         | unix=$rtcUnix"
    Write-Host "Écart : $diff s (tolérance = $TimeToleranceSeconds s)"

    if ($diff -le $TimeToleranceSeconds) {
        Write-Host "OK : date/heure RTC cohérente."
        return
    }

    $ans = Read-Host "Mettre à jour le RTC du PI avec la date+heure du PC ? (o/n)"
    if ($ans -notin @("o", "O", "oui", "OUI", "y", "Y", "yes", "YES")) {
        Write-Host "Mise à l'heure annulée."
        return
    }

    # Écriture RTC (pas besoin de sudo si l'utilisateur a accès I2C, comme en manuel)
    Set-RtcFromPcNow -pcNow (Get-Date)

    # Relire RTC après écriture
    $rtcHuman2 = Get-PiHumanTime
    $rtcUnix2 = Get-PiUnixTime
    $diff2 = [math]::Abs($pcUnix - $rtcUnix2)

    Write-Host "PI (RTC après) : $rtcHuman2 | unix=$rtcUnix2"
    Write-Host "Nouvel écart : $diff2 s"
    Write-Host "Important : le PI se resynchronise sur le RTC au démarrage. Un reboot est nécessaire pour que l'heure système suive."

    $ans2 = Read-Host "Rebooter le PI maintenant ? (o/n)"
    if ($ans2 -in @("o", "O", "oui", "OUI", "y", "Y", "yes", "YES")) {
        $rSudo = Invoke-RemoteShell "sudo -n true"
        if ($rSudo.ExitCode -ne 0) {
            throw "Impossible : sudo nécessite un mot de passe (ou non autorisé) pour l'utilisateur '$UserName'."
        }
        $rRb = Invoke-RemoteShell "sudo -n reboot"
        if ($rRb.ExitCode -ne 0) {
            throw "Échec reboot (code $($rRb.ExitCode)). Sortie:`n$($rRb.Output)"
        }
        Write-Host "Reboot demandé."
    }
    else {
        Write-Host "Reboot non effectué (à faire plus tard)."
    }
}

# ======================================================
# Comptage fiable (utilisé par option 3/4)
# ======================================================
function Get-RemoteLogCount {
    $r = Invoke-RemoteShell "cd '$RemoteLogsDir' && echo __COUNT__ && ls -1 $FileMask 2>/dev/null | wc -l"
    if ($r.ExitCode -ne 0) { throw "Erreur comptage fichiers sur le Pi (code $($r.ExitCode)).`n$($r.Output)" }
    if ($r.Output -match '__COUNT__\s*[\r\n]+(?<n>\d+)\s*') { return [int]$matches['n'] }
    return 0
}

# ======================================================
# 2) SCAN
# ======================================================
function Scan-Logs {
    $cmds = @(
        "cd '$RemoteLogsDir' && echo __COUNT__ && ls -1 $FileMask 2>/dev/null | wc -l",
        "cd '$RemoteLogsDir' && echo __DIRS__  && ls -1d */ 2>/dev/null || true",
        "cd '$RemoteLogsDir' && echo __FILES__ && ls -1 $FileMask 2>/dev/null || true"
    )

    $r = Invoke-RemoteShellMulti $cmds
    if ($r.ExitCode -ne 0) { throw "Scan échoué (code $($r.ExitCode)). Voir: $WinSCPLog`n$($r.Output)" }

    $text = $r.Output

    $count = 0
    if ($text -match '__COUNT__\s*[\r\n]+(?<n>\d+)\s*') { $count = [int]$matches['n'] }

    $dirs = @()
    if ($text -match '__DIRS__') {
        $dirsPart = ($text -split '__DIRS__', 2)[1]
        $dirsPart = ($dirsPart -split '__FILES__', 2)[0]
        $dirs = @(
            ($dirsPart -split "`r?`n") |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
        )
    }

    $files = @()
    if ($text -match '__FILES__') {
        $filesPart = ($text -split '__FILES__', 2)[1]
        $files = @(
            ($filesPart -split "`r?`n") |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
        )
    }

    Write-Host "===== SCAN LOGS (Pi) ====="
    Write-Host "Répertoire : $RemoteLogsDir"
    Write-Host "Masque     : $FileMask"
    Write-Host "Fichiers (niveau racine) : $count"
    Write-Host ""

    Write-Host "Sous-répertoires (archives) :"
    if (@($dirs).Count -eq 0) { Write-Host " - (aucun)" }
    else { $dirs | ForEach-Object { Write-Host " - $_" } }

    Write-Host ""
    Write-Host "Fichiers logs (niveau racine) :"
    if (@($files).Count -eq 0) { Write-Host " - (aucun)" }
    else { $files | ForEach-Object { Write-Host " - $_" } }

    Write-Host "==========================="
}

# ======================================================
# 3) TELECHARGEMENT (sync + MD5)
# ======================================================
function Get-LocalSnapshot {
    param(
        [Parameter(Mandatory)] [string]$dir,
        [Parameter(Mandatory)] [string]$pattern
    )

    $map = @{}
    Get-ChildItem -LiteralPath $dir -File -Filter $pattern -ErrorAction SilentlyContinue | ForEach-Object {
        $map[$_.Name] = [pscustomobject]@{
            Name = $_.Name
            Size = $_.Length
            Time = $_.LastWriteTimeUtc
        }
    }
    return $map
}

function Download-Logs {
    $count = Get-RemoteLogCount
    Write-Host "Téléchargement : fichiers présents sur le Pi (niveau racine) = $count"
    if ($count -eq 0) {
        Write-Host "Aucun fichier à télécharger (probablement déjà archivés)."
        return
    }

    # --- Snapshot local avant (strict sur .txt)
    $before = Get-LocalSnapshot -dir $LocalDir -pattern "ELISA_Prod_Log_*.txt"

    # --- Étape 1 : Sync rapide (nouveaux + modifiés taille/date)
    $syncScript = @"
option batch abort
option confirm off
$(New-OpenLine)
cd "$RemoteLogsDir"
synchronize local -filemask="$FileMask" "$LocalDir"
exit
"@
    Write-Host "[Téléchargement] Étape 1 : synchronisation..."
    $r1 = Invoke-WinSCP $syncScript
    if ($r1.ExitCode -ne 0) {
        throw "WinSCP synchronize a échoué (code $($r1.ExitCode)). Voir: $WinSCPLog`nSortie:`n$($r1.Output)"
    }

    # --- Snapshot local après
    $after = Get-LocalSnapshot -dir $LocalDir -pattern "ELISA_Prod_Log_*.txt"

    # --- Détection "nouveaux" uniquement (on n'affiche pas les 'mis à jour' ici)
    $newFiles = New-Object System.Collections.Generic.List[string]
    foreach ($kv in $after.GetEnumerator()) {
        $name = $kv.Key
        if (-not $before.ContainsKey($name)) {
            $newFiles.Add($name)
        }
    }

    Write-Host ""
    Write-Host "Résumé :"
    Write-Host (" - Nouveaux téléchargés : {0}" -f $newFiles.Count)
    if ($newFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "Nouveaux fichiers :"
        $newFiles | Sort-Object | ForEach-Object { Write-Host " - $_" }
    }
    else {
        Write-Host " - Aucun nouveau fichier."
    }

    # --- Étape 2 : Contrôle contenu (MD5) + retéléchargement si différent
    Write-Host ""
    Write-Host "[Téléchargement] Étape 2 : contrôle contenu (MD5) + retéléchargement si nécessaire..."

    $redownloadCount = 0

    $r2 = Invoke-RemoteShell "cd '$RemoteLogsDir' && md5sum $FileMask 2>/dev/null || true"
    if ($r2.ExitCode -ne 0) { throw "Lecture MD5 distante échouée (code $($r2.ExitCode)).`n$($r2.Output)" }

    $remote = @{}
    foreach ($line in ($r2.Output -split "`r?`n")) {
        if ($line -match '^(?<md5>[a-f0-9]{32})\s+(?<name>.+)$') {
            $remote[$matches['name'].Trim()] = $matches['md5'].ToLowerInvariant()
        }
    }

    if ($remote.Count -eq 0) {
        Write-Host "Aucun MD5 récupéré (aucun fichier correspondant au masque)."
        Write-Host ""
        Write-Host "Bilan téléchargement :"
        Write-Host (" - Nouveaux fichiers : {0}" -f $newFiles.Count)
        Write-Host (" - Fichiers corrigés (MD5 différent) : {0}" -f 0)
        return
    }

    $toRedownload = New-Object System.Collections.Generic.List[string]
    foreach ($kv in $remote.GetEnumerator()) {
        $name = $kv.Key
        $remoteMd5 = $kv.Value

        $localPath = Join-Path $LocalDir $name
        if (-not (Test-Path -LiteralPath $localPath)) {
            # absent localement => on le récupère
            $toRedownload.Add($name)
            continue
        }

        $localMd5 = (Get-FileHash -LiteralPath $localPath -Algorithm MD5).Hash.ToLowerInvariant()
        if ($localMd5 -ne $remoteMd5) {
            $toRedownload.Add($name)
        }
    }

    if ($toRedownload.Count -eq 0) {
        Write-Host "OK : aucun fichier différent détecté au niveau contenu."
        Write-Host ""
        Write-Host "Bilan téléchargement :"
        Write-Host (" - Nouveaux fichiers : {0}" -f $newFiles.Count)
        Write-Host (" - Fichiers corrigés (MD5 différent) : {0}" -f 0)
        return
    }

    $redownloadCount = $toRedownload.Count

    Write-Host ("Fichiers à retélécharger (contenu différent) : {0}" -f $toRedownload.Count)
    $toRedownload | Sort-Object | ForEach-Object { Write-Host " - $_" }

    # Pas de -overwrite sur certaines versions WinSCP -> supprimer avant
    foreach ($name in $toRedownload) {
        $localPath = Join-Path $LocalDir $name
        if (Test-Path -LiteralPath $localPath) { Remove-Item -LiteralPath $localPath -Force }
    }

    $getLines = ($toRedownload | ForEach-Object {
            $dst = Join-Path $LocalDir $_
            'get -transfer=binary -resume "{0}" "{1}"' -f $_, $dst
        }) -join "`r`n"

    $redlScript = @"
option batch abort
option confirm off
$(New-OpenLine)
cd "$RemoteLogsDir"
$getLines
exit
"@
    $r3 = Invoke-WinSCP $redlScript
    if ($r3.ExitCode -ne 0) {
        throw "Retéléchargement échoué (code $($r3.ExitCode)). Voir: $WinSCPLog`nSortie:`n$($r3.Output)"
    }

    Write-Host "Retéléchargement terminé."

    Write-Host ""
    Write-Host "Bilan téléchargement :"
    Write-Host (" - Nouveaux fichiers : {0}" -f $newFiles.Count)
    Write-Host (" - Fichiers corrigés (MD5 différent) : {0}" -f $redownloadCount)
}

# ======================================================
# 4) ARCHIVAGE
# ======================================================
function Archive-Logs {
    $count = Get-RemoteLogCount
    Write-Host "Archivage : fichiers présents sur le Pi (niveau racine) = $count"
    if ($count -eq 0) {
        Write-Host "Aucun fichier à archiver."
        return
    }

    $archiveName = Read-Host "Nom du sous-répertoire d'archive à créer dans logs (ex: archive_2026-01-13_0815)"
    if ([string]::IsNullOrWhiteSpace($archiveName)) { Write-Host "Archivage annulé (nom vide)."; return }
    if ($archiveName -match '[\\/:*?"<>|]') { Write-Host 'Nom invalide (caractères interdits: \ / : * ? " < > |).'; return }
    if ($archiveName -match '^\.+$') { Write-Host "Nom invalide."; return }

    $archiveDir = "$RemoteLogsDir/$archiveName"
    Write-Host "Archive cible : $archiveDir"

    $cmds = @(
        "mkdir -p '$archiveDir'",
        "cd '$RemoteLogsDir' && mv -f $FileMask '$archiveDir/' 2>/dev/null || true",
        "cd '$RemoteLogsDir' && echo __COUNT__ && ls -1 $FileMask 2>/dev/null | wc -l"
    )

    $r = Invoke-RemoteShellMulti $cmds
    if ($r.ExitCode -ne 0) { throw "Archivage échoué (code $($r.ExitCode)). Voir: $WinSCPLog`n$($r.Output)" }

    $after = 0
    if ($r.Output -match '__COUNT__\s*[\r\n]+(?<n>\d+)\s*') { $after = [int]$matches['n'] }
    Write-Host "Archivage terminé. Fichiers restants (masque) dans logs : $after"
}

# ======================================================
# MAIN
# ======================================================
try {
    Write-Host "Script : $PSCommandPath"
    Write-Host "Log PS     : $PsLog"
    Write-Host "Log WinSCP : $WinSCPLog"

    :MainLoop while ($true) {
        Show-Menu
        $choice = (Read-Host "Choix").Trim()

        switch ($choice) {
            "1" {
                Show-ActionHeader "VÉRIFICATION HORODATAGE (PC vs PI)"
                Check-Time-And-OfferSync
                Show-ActionFooter
            }
            "2" {
                Show-ActionHeader "SCAN DES LOGS (compte + archives + liste)"
                Scan-Logs
                Show-ActionFooter
            }
            "3" {
                Show-ActionHeader "TÉLÉCHARGEMENT DES LOGS"
                Download-Logs
                Show-ActionFooter
            }
            "4" {
                Show-ActionHeader "ARCHIVAGE DES LOGS SUR LE PI"
                Archive-Logs
                Show-ActionFooter
            }
            "0" {
                Show-ActionHeader "SORTIE"
                Write-Host "Fermeture du script..."
                break MainLoop
            }
            default {
                Write-Host ""
                Write-Host "Choix invalide."
                [void](Read-Host "Appuyer sur Entrée pour revenir au menu")
            }
        }
    }
}
finally {
    Stop-Transcript | Out-Null
}

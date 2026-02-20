function Show-Menu {
    param (
        [string]$Title,
        [array]$Options
    )

    $Host.UI.RawUI.WindowTitle = $Title
    $selectedIndex = 0

    while ($true) {
        Clear-Host
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "  $Title" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "Utilisez les flèches HAUT/BAS et ENTRÉE pour valider" -ForegroundColor Gray
        Write-Host "Appuyez sur ÉCHAP pour quitter.`n" -ForegroundColor Gray

        for ($i = 0; $i -lt $Options.Count; $i++) {
            if ($i -eq $selectedIndex) {
                Write-Host " > $($Options[$i].Name) " -ForegroundColor Black -BackgroundColor Cyan
            }
            else {
                Write-Host "   $($Options[$i].Name) "
            }
        }

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        
        switch ($key.VirtualKeyCode) {
            38 {
                # Flèche Haut
                $selectedIndex--
                if ($selectedIndex -lt 0) { $selectedIndex = $Options.Count - 1 }
            }
            40 {
                # Flèche Bas
                $selectedIndex++
                if ($selectedIndex -ge $Options.Count) { $selectedIndex = 0 }
            }
            13 {
                # Entrée
                return $Options[$selectedIndex]
            }
            27 {
                # Échap
                return $null
            }
        }
    }
}

$scripts = @(Get-ChildItem -Path $PSScriptRoot -Filter "*.ps1" | Where-Object Name -NotMatch "^Option")

if ($scripts.Count -eq 0) {
    Write-Warning "Aucun script PowerShell (autre que les lanceurs) trouvé."
    Pause
    return
}

$selection = Show-Menu -Title "LANCEUR DE SCRIPTS POWERSHELL" -Options $scripts

if ($selection) {
    Clear-Host
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "   EXÉCUTION DE : $($selection.Name)" -ForegroundColor Cyan
    Write-Host "============================================`n" -ForegroundColor Cyan
    
    # Exécution dans un nouveau processus isolé
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($selection.FullName)`"" -Wait -NoNewWindow
    
    Write-Host "`n============================================" -ForegroundColor Cyan
    Write-Host "   Exécution terminée." -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
}

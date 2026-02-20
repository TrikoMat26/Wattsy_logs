$scripts = Get-ChildItem -Path $PSScriptRoot -Filter "*.ps1" | Where-Object Name -NotMatch "^Option"

if ($scripts.Count -eq 0) {
    Write-Warning "Aucun script PowerShell (autre que les lanceurs) trouvé."
    Pause
    return
}

# Ouvre une fenêtre avec barre de recheche
$selection = $scripts | Select-Object Name, LastWriteTime | Out-GridView -Title "Sélectionnez un script PowerShell à exécuter" -PassThru

if ($selection) {
    $scriptToRun = ($scripts | Where-Object Name -eq $selection.Name).FullName
    Clear-Host
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "   EXÉCUTION DE : $($selection.Name)" -ForegroundColor Cyan
    Write-Host "============================================`n" -ForegroundColor Cyan
    
    # Exécution dans un nouveau processus pour éviter les conflits de DPI/Scaling avec Out-GridView (WPF vs WinForms)
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptToRun`"" -Wait -NoNewWindow
    
    Write-Host "`n============================================" -ForegroundColor Cyan
    Write-Host "   Exécution terminée." -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
}

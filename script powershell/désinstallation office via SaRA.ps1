# Ce script désinstalle Microsoft Office 2016 à l'aide du scénario OfficeScrubScenario via SaRAcmd.exe

$SaRAcmd = "$PSScriptRoot\SaRAcmd.exe"

# Vérifie que l'exécutable existe à côté du script
if (-Not (Test-Path $SaRAcmd)) {
    Write-Error " SaRAcmd.exe introuvable dans le dossier du script. Place-le au même endroit que ce script."
    exit 1
}

# Lancer le scénario OfficeScrub qui désinstalle Office 2016
try {
    Write-Output " Lancement de la désinstallation de Microsoft Office 2016 via SaRA..."
    Start-Process -FilePath $SaRAcmd `
                  -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion 2016" `
                  -Wait -NoNewWindow
    Write-Output " Scénario terminé."
}
catch {
    Write-Error " Une erreur est survenue lors de l'exécution de SaRAcmd.exe : $_"
}

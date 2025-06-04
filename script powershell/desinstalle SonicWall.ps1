# Définir le chemin complet du fichier MSI (à adapter si nécessaire)
$msiPath = "$PSScriptRoot\GVCInstall64.msi"

# Vérifier si le fichier existe
if (Test-Path $msiPath) {
    Write-Host "Fichier MSI trouvé. Démarrage de la désinstallation silencieuse..."
    Start-Process "msiexec.exe" -ArgumentList "/x `"$msiPath`" /quiet /norestart" -Wait
    Write-Host "Désinstallation terminée."
} else {
    Write-Host "Erreur : le fichier MSI n'a pas été trouvé à l'emplacement suivant : $msiPath"
}

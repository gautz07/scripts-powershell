# Ce script désinstalle Microsoft Office 2016 (msi) à l'aide du scénario OfficeScrubScenario via SaRAcmd.exe
# Pour que ce script fonctionne, il doit etre placer dans le meme répertoire que l'executable SaRAcmd.exe
Function rechercheOffice2016 ()
{
    # Recherche Office 2016 installé via Windows Installer (MSI)
    $officeEntries = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
                                     HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,
                                     HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\* |
        Where-Object { $_.DisplayName -like "*Microsoft Office*2016*"}
    if ($officeEntries) {return $true}
    else {return $false}
}

if (rechercheOffice2016 -eq $true) {

    $SaRAcmd = "$PSScriptRoot\SaRAcmd.exe"

    # Vérifie que l'exécutable existe à côté du script
    if (-Not (Test-Path $SaRAcmd)) {
        Write-Error " SaRAcmd.exe introuvable dans le dossier du script. Placez-le au même endroit que ce script."
        exit 1
    }

    # Lancer le scénario OfficeScrub qui désinstalle Office 2016 de maniere automatique

    Write-Output " Lancement de la désinstallation de Microsoft Office 2016 via SaRA..."
    $processInfo = new-Object System.Diagnostics.ProcessStartInfo($SaRAcmd);
    $processInfo.Arguments = "-s OfficeScrubScenario -AcceptEULA -OfficeVersion 2016";
    $processInfo.CreateNoWindow = $true;     
    $processInfo.UseShellExecute = $false;
    $processInfo.RedirectStandardOutput = $true;

    $process = [System.Diagnostics.Process]::Start($processInfo);
    $process.StandardOutput.ReadToEnd();     # Affiche la sortie de SaRA CMD dans la fenêtre PowerShell
    $process.WaitForExit();
    $success = $process.HasExited -and $process.ExitCode -eq 0; #vérifie que SaRAcmd s'est bien fermé normalement sans erreurs 
    $process.Dispose();

    if ($success -eq $false) {
        Write-Error "Une erreur est survenue lors de l'exécution de SaRAcmd.exe. Code de sortie : $($process.ExitCode)"

    }

    # Vérification après désinstallation 
    if (rechercheOffice2016 -eq $true) {
        Write-Output "Office 2016 est toujours présent. Désinstallation échouée."
    } else {
        Write-Output "Office 2016 a été désinstallé avec succès."
    }
}else {
    Write-Output "Aucun Office 2016 (MSI) trouvé sur cette machine." 
    }
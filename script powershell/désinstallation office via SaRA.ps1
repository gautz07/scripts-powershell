# Ce script désinstalle Microsoft Office 2016 à l'aide du scénario OfficeScrubScenario via SaRAcmd.exe
Function verificationOffice365() {
    $paths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    )

    $existingPaths = @()
    foreach ($path in $paths) {
        $basePath = ($path -split '\\\*')[0]
        if (Test-Path $basePath) {
            $existingPaths += $path
        }
    }

    $officeEntries = foreach ($path in $existingPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    }

    $officeEntries = $officeEntries | Where-Object {
        $_.DisplayName -match "Microsoft 365|Office 365|Office ProPlus|Microsoft Office.*365"
    }
    
    if ($officeEntries) {return $true}
    else {return $false}
}


if (rechercheOffice2016 -eq $true) {

    $SaRAcmd = "$PSScriptRoot\SaRACmd\SaRAcmd.exe"

    # Vérifie que l'exécutable existe à côté du script
    if (-Not (Test-Path $SaRAcmd)) {
        Write-Error " SaRAcmd.exe introuvable."
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
    Write-Output "Aucun Office 2016 trouvé sur cette machine." 
    }

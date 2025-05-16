# Ce script désinstalle Microsoft Office 2016 à l'aide du scénario OfficeScrubScenario via SaRAcmd.exe
# Pour que ce script fonctionne, il doit etre placer dans le meme répertoire que l'executable SaRAcmd.exe
Function rechercheOffice2016 ()
{
$paths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
)

$existingPaths = @()
foreach ($path in $paths) {
    $basePath = ($path -split '\\\*')[0]  # enlève le wildcard pour test
    if (Test-Path $basePath) {
        $existingPaths += $path
    }
}

$officeEntries = foreach ($path in $existingPaths) {
    Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
}

$officeEntries = $officeEntries | Where-Object { $_.DisplayName -like "*Microsoft Office*2016*" }

    if ($officeEntries) {return $true}
    else {return $false}
}

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

#-------installation de office 365 -------
$odtExe = "$PSScriptRoot\office 365\officedeploymenttool_18730-20142.exe"
$configXmlPath = "$PSScriptRoot\office 365\Configuration.xml"

Write-Host "`$PSScriptRoot = $PSScriptRoot"
Write-Host "`$odtExe = $odtExe"
Test-Path $odtExe
# Extraire le contenu de l'exe
Write-Host "Extraction de l'Office Deployment Tool..."
Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:`"$PSScriptRoot`"" -Wait -NoNewWindow
if (Test-Path "$PSScriptRoot\setup.exe") {
    Write-Host " setup.exe trouvé."
} else {
    Write-Error " setup.exe introuvable dans $PSScriptRoot après extraction."
    exit 1
}


# Copier le fichier Configuration.xml fourni par l'utilisateur
Write-Host "Copie du fichier Configuration.xml..."
$xmlContent = @"
<Configuration ID="5e9262a5-ddfb-4bc0-806e-142dde0463c7">
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365ProPlusRetail">
      <Language ID="fr-fr" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
$xmlContent | Set-Content -Encoding UTF8 -Path $configXmlPath

# Lancer l'installation
Write-Host "Lancement de l'installation d'Office..."
Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure `"$configXmlPath`"" -Wait -NoNewWindow

Write-Host " Installation terminée."

# Vérification après installation
if (verificationOffice365 -eq $true) {
    Write-Output " Microsoft 365 a été installé avec succès."
} else {
    Write-Error " Échec : Microsoft 365 n'a pas été détecté après l'installation."
}
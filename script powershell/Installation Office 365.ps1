Function verificationOffice365() {
    #fonction qui recherche des installations d'office 365
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

$odtExe = "$PSScriptRoot\officedeploymenttool_18730-20142.exe"
$configXmlPath = "$PSScriptRoot\Configuration.xml"

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


# Copier le fichier Configuration.xml
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

# Lanceement de l'installation
Write-Host "Lancement de l'installation d'Office..."
Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure Configuration.xml" -Wait -NoNewWindow

Write-Host " Installation terminée."

# Vérification après installation
if (verificationOffice365 -eq $true) {
    Write-Output " Microsoft 365 a été installé avec succès."
} else {
    Write-Error " Échec : Microsoft 365 n'a pas été détecté après l'installation."
}
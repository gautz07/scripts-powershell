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
Start-Process -FilePath "$PSScriptRoot\setup.exe" -ArgumentList "/configure Configuration.xml" -Wait -NoNewWindow

Write-Host " Installation terminée."
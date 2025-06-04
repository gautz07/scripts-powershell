#pour que le script fonctionne il faut que le dossier "Sage Safe X3 Client 117" pour l'installation de SAGE 
function detection_du_.NET {
    # Clé du registre pour .NET Framework 4.x
    $regKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    $dismOutput = DISM /Online /Get-Features | Select-String "NetFx3"

    if (Test-Path $regKey) {
        $release = (Get-ItemProperty -Path $regKey).Release #test la detection du .NET version 4.x
        if ($release) {
            Write-Host ".NET Framework version 4.x detectée"
            return $true
        } else {
            Write-Host "Aucun .NET Framework 4.x détecté."
        }
    } elseif($dismOutput -match "NetFx3\s*:\s*Enabled") { #test la detection du .NET version 3.5
        Write-Host ".NET Framework 3.5 détecté."
        return $true
    }else{
        return $false
    }
}

$net = detection_du_.NET
if ($net -eq $false) {
     Write-Host ".NET Framework 3.5 n'est pas installé."

    # Exécuter DISM pour installer le .NET 3.5
    $sourcePath = "$PSScriptRoot\sxs"
    $cmd = "DISM /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:$sourcePath"

    Write-Host "Exécution de : $cmd"
    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $cmd -Wait -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host ".NET Framework installé avec succès !" -ForegroundColor Green
    } else {
        Write-Host "Échec de l'installation de .NET Framework (code : $($process.ExitCode))" -ForegroundColor Red
    }

}
$issFile = "$PSScriptRoot\Sage Safe X3 Client 117\unattended.iss"

Write-Output "lancement de l'installation de sage safe X3"
Start-Process -FilePath "$PSScriptRoot\Sage Safe X3 Client 117\Install.cmd" -Wait -NoNewWindow # lancement du fichier CMD qui fait l'installation silencieuse 
$installPath = Select-String -Path $issFile -Pattern 'szPath=' | ForEach-Object {
    ($_ -split '=')[1].Trim()
}
Write-Host "Chemin d'installation : $installPath"

Copy-Item -Path "$($PSScriptRoot)\Sage Safe X3 Client 117\Spécifique CDP Fareva\DLL a remplacer\X3scales.dll" -Destination $installPath -Force
Copy-Item -Path "$($PSScriptRoot)\Sage Safe X3 Client 117\Spécifique CDP Fareva\DLL a remplacer\X3ScaFRA.dbm" -Destination "$($installPath)\lan" -Force
Copy-Item -Path "$($PSScriptRoot)\Sage Safe X3 Client 117\Spécifique CDP Fareva\Pictogrames a ajouter" -Destination "$($installPath)\\Icons\Risquesecu_jpg" -Recurse -Force

# Chemin vers le fichier .ttf
$FontSource = "$($PSScriptRoot)\Sage Safe X3 Client 117\Spécifique CDP Fareva\Fonts\C39T36L.ttf"
$FontName = "C39T36L.ttf"  # Nom du fichier

Write-Host "installation de la police 'C39T36L'."
# Dossier des polices système
$FontsDir = "$env:WINDIR\Fonts"
$FontDest = Join-Path -Path $FontsDir -ChildPath $FontName

# Copier la police dans le dossier Fonts (si pas déjà là)
If (!(Test-Path $FontDest)) {
    Copy-Item -Path $FontSource -Destination $FontDest
}

# Enregistrement dans le Registre (nécessite admin)
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
$FontRegName = "C39T36L (TrueType)"  # Nom affiché dans la liste de polices
Set-ItemProperty -Path $RegPath -Name $FontRegName -Value $FontName

Write-Host "Police installée avec succès."

Write-Output "installation de sage safe X3 terminée"
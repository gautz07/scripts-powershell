#ce script doit etre placé dans le meme repertoire que le fichier iso de windows 
#et que le dossier Sage Safe X3 Client 117 pou l'installation de SAGE 
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

$dismOutput = DISM /Online /Get-Features | Select-String "NetFx3"
$net = detection_du_.NET
if ($net -eq $false) {
     Write-Host ".NET Framework 3.5 n'est pas installé. Montage de l'ISO en cours..."
    # Trouve le fichier ISO dans le même dossier que le script
    $isoPath = Get-ChildItem -Path $PSScriptRoot -Filter *.iso | Select-Object -First 1

    if (-not $isoPath) {
        Write-Host "Aucun fichier ISO trouvé dans le dossier du script." -ForegroundColor Red
        exit 1
    }

    # Monte l'image ISO
    $mount = Mount-DiskImage -ImagePath $isoPath.FullName -PassThru
    Start-Sleep -Seconds 2 # pause pour s'assurer du montage

    # Obtenir la lettre du lecteur monté
    $volume = Get-Volume -DiskImage $mount
    $driveLetter = $volume.DriveLetter

    Write-Host "ISO monté sur le lecteur $driveLetter`:"

    # Exécuter DISM pour installer le .NET
    $sourcePath = "$driveLetter`:\sources\sxs"
    $cmd = "DISM /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:$sourcePath"

    Write-Host "Exécution de : $cmd"
    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $cmd -Wait -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host ".NET Framework 3.5 installé avec succès !" -ForegroundColor Green
    } else {
        Write-Host "Échec de l'installation de .NET Framework 3.5 (code : $($process.ExitCode))" -ForegroundColor Red
    }

    # Nettoyage : démonter l'image
    Dismount-DiskImage -ImagePath $isoPath.FullName
    Write-Host "ISO démontée."
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

Write-Output "installation de sage safe X3 terminée"
# Lancer en tant qu'administrateur

# Fonction de désinstallation propre
function Uninstall-App {
    param (
        [string]$app
    )

    $unistalls = Get-AppxPackage -AllUsers | Where-Object { $_.Name -like "*$app*" }
    $prov = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$app*" }

    if ($unistalls) {
        foreach ($unistall in $unistalls) {
            Write-Output "Désinstallation : $($unistall.Name)"
            Try {
                Remove-AppxPackage -Package $unistall.PackageFullName -AllUsers -ErrorAction Stop
            } Catch {
                Write-Warning "Échec de la désinstallation : $($unistall.Name)"
            }
        }
    }

    if ($prov) {
        foreach ($p in $prov) {
            Write-Output "Suppression du provisioning : $($p.DisplayName)"
            Try {
                Remove-AppxProvisionedPackage -Online -PackageName $p.PackageName -ErrorAction Stop
            } Catch {
                Write-Warning "Échec de la suppression du provisioning : $($p.DisplayName)"
            }
        }
    }
}

# Liste des applications désinstallables
$appsToUninstall = @(
    "windowsalarms",
    "windowsmaps",
    "WebMediaExtensions",
    "HEIFImageExtension",
    "WebpImageExtension",
    "zunevideo",
    "zunemusic",
    "bingweather",
    "solitairecollection",
    "skypeapp",
    "xboxapp",
    "Xbox.TCUI",
    "XboxGameOverlay",
    "XboxGamingOverlay",
    "XboxSpeechToTextOverlay"
)

foreach ($app in $appsToUninstall) {
    Write-Output "`nRecherche et désinstallation de : $app"
    Uninstall-App -app $app
}

# Désactivation via registre des apps non désinstallables
Write-Output "`nDésactivation de Cortana via registre"
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0

Write-Output "Désactivation de PeopleExperienceHost (Contacts)"
New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\People" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\People" -Name "PeopleBand" -Value 0

Write-Output "Désactivation des fonctions Xbox via registre"
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR" -Name "AllowGameDVR" -Value 0

New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR" -Name "AppCaptureEnabled" -Value 0

New-Item -Path "HKCU:\Software\Microsoft\GameBar" -Force | Out-Null
Set-ItemProperty -Path "HKCU:\Software\Microsoft\GameBar" -Name "UseNexusForGameBarEnabled" -Value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\GameBar" -Name "ShowGameBarTips" -Value 0

Write-Output "`nNettoyage terminé."

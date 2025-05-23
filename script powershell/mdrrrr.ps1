function detection_du_.NET {
    $regKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    $dismOutput = DISM /Online /Get-Features | Select-String "NetFx3"

    if (Test-Path $regKey) {
        $release = (Get-ItemProperty -Path $regKey).Release
        
        if ($release -and $release -gt 0) {
            Write-Host ".NET Framework version 4.x detectee"
            return $true
        } else {
            Write-Host "Clé présente mais aucune version 4.x valide detectee"
            return $false
        }
    }

    elseif ($dismOutput -match "NetFx3\s*:\s*Enabled") {
        Write-Host ".NET Framework 3.5 detecte."
        return $true
    }
    else{
        Write-Host "Aucun .NET Framework detecte"
        return $false
    }   
}

$test = detection_du_.NET
write-output $test

if (detection_du_.NET -eq $False){
    if (detection_du_.NET -eq $true){
        write-output "mdrrrrrrrrr"
    }
    
}else{
    write-output "c'est bon"
}


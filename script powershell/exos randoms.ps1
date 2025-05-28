########## exercice 1 Script de nettoyage intelligent ##########
function NettoyageFichier {
    param (
        [string]$path
    )
    $today = Get-Date
    foreach ($dossier in Get-ChildItem -path $path) {
        $ecart = $today - $dossier.lastWriteTime
        if ($ecart.Days -gt 30 -and $dossier.Name -notlike "*important*") {
            
            Remove-Item $dossier.FullName
        }
    }   
}
NettoyageFichier -path  'C:\Users\BARDEGA1\Documents'

########## exercice 2 Script de "fiche utilisateur" ##########
function fiche_user {
    param (
        [string]$path
    )
    Write-Output 'saisissez votre nom'
    $nom = Read-Host
    Write-Output 'saisissez votre premon'
    $prenom = Read-Host
    Write-Output 'saisissez votre adresse Email'
    $email = Read-Host

    $datas = @($nom,$prenom,$email)
    Write-Output "bonjour $($prenom) nous allons vous enregistrer"

    $fichier = Get-Content -path $path 
    $i = 0
    $nouveau_fichier = @()
    foreach($ligne in $fichier){
        if ($ligne -match 'Nom|Prenom|Email'){
            $ligne += $datas[$i]
            $i += 1
        }
        $nouveau_fichier += $ligne  
    }
    Write-Output $nouveau_fichier
    Set-Content -Path $path -Value $nouveau_fichier
}
#fiche_user -path  'C:\Users\BARDEGA1\Documents\fiche utilisateur important.txt'

function menu_actions {
    param (
        [string]$path
    )
    $choix_user = ''

    while ($choix_user -ne '4'){
  Write-Host "===== MENU ====="
        Write-Host "1. Lister les fichiers du dossier courant"
        Write-Host "2. Supprimer tous les fichiers .log du dossier courant"
        Write-Host "3. Afficher l adresse IP locale"
        Write-Host "4. Quitter"
        Write-Host "================"
        $choix_user = Read-Host "Votre choix"

        switch ($choix_user) {
            '1' {
                Write-Host "Fichiers dans $path :"
                Get-ChildItem -Path $path
            }
            '2' {
                Write-Host "Suppression des fichiers .log dans $path"
                Get-ChildItem -Path $path -Filter *.log | Remove-Item -Force
            }
            '3' {
                $ip = (ipconfig | Select-String "IPv4").ToString().Split(":")[1].Trim()
                Write-Host "Adresse IP locale : $ip"
            }
            '4' {
                Write-Host "Fermeture du menu..."
            }
            default {
                Write-Host "Choix invalide, veuillez reessayer."
            }
        }
        Read-Host "Appuyez sur Enter..."
    }
}
menu_actions -path  'C:\Users\BARDEGA1\Documents'
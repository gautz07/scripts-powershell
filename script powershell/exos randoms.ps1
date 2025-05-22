########## exercice 1 Script de nettoyage intelligent ##########
function NettoyageFichier {
    param (
        [string]$path
    )
    $today = Get-Date
    foreach ($dossier in Get-ChildItem -path $path) {
        $ecart = $today - $dossier.lastWriteTime
        if ($ecart.Days -gt 30 -and $dossier.Name -notlike "*important*") {
            
            Remove-Item $dossier
        }
    }   
}
NettoyageFichier -path  C:\Users\BARDEGA1\Documents

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

    Write-Output "bonjour $($prenom) nous allons vous enregistrer"

    $fichier = Get-Content -path $path 
}
fiche_user -path C:\Users\BARDEGA1\Documents
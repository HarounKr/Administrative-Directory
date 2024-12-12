#Lister les adresses IP utilisees par des utilisateurs 

$events = Get-WinEvent -LogName Security -MaxEvents 20000 | Where-Object {
    $_.Id -eq 4624
} | Select-Object TimeCreated, @{Name="Utilisateur";Expression={$_.Properties[5].Value}}, 
    @{Name="Adresse IP";Expression={$_.Properties[18].Value}}

$outputFile = "$PWD\Rapport_Evenements_usersIP.xlsx"

if ($events){
# Exporter vers un fichier Excel
    $events | Export-Excel -Path $outputFile -WorksheetName 'Événements' -AutoSize

    Write-Host "Les événements ont été exportés dans le fichier $outputFile"
} else {
    Write-Host "Aucun événement trouvé"

}
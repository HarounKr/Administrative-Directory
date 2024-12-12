# Lister les incidents de securite

$events = Get-WinEvent -LogName Security | Where-Object {
    $_.Id -in 4625,4647,4720,4723,4740,4768,4776,4794
} | Select-Object TimeCreated, Id, @{Name="Type Incident";Expression={
    switch ($_.Id) {
        4625 {"Échec Connexion"}
        4647 {"Déconnexion"}
        4720 {"Création Compte"}
        4723 {"Échec Modification Mot de Passe"}
        4740 {"Compte Verrouillé"}
        4768 {"Ticket Kerberos"}
        4776 {"Authentification NTLM"}
        4794 {"Changement de Service"}
    }
}}, Message

$outputFile = "$PWD\Rapport_Evenements_SecurityIncidents.xlsx"

if ($events){
# Exporter vers un fichier Excel
    $events | Export-Excel -Path $outputFile -WorksheetName 'Événements' -AutoSize

    Write-Host "Les événements ont été exportés dans le fichier $outputFile"
} else {
    Write-Host "Aucun événement trouvé"

}
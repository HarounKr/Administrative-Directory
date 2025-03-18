# Afficher les événements concernant un utilisateur entre deux heures spécifiques

param(
    [Parameter(Mandatory=$true)]    
    [string]$User,       # Nom de l'utilisateur
    [Parameter(Mandatory=$true)]
    [string]$StartTime,  
    [Parameter(Mandatory=$true)]
    [string]$EndTime 
)

try {
    $startTime = Get-Date $StartTime -Format "dd/MM/yyyy HH:mm:ss" -ErrorAction Stop
    $startTime
    $endTime = Get-Date $EndTime -Format "dd/MM/yyyy HH:mm:ss" -ErrorAction Stop
    $endTime
} catch {
    Write-Host "Les dates fournies ne sont pas valides. "
    exit 1
}

$User = $User.ToLower()

$events= Get-WinEvent -FilterHashtable @{
    LogName = 'Security';
    StartTime = $startTime;
    EndTime = $endTime;
} |  Where-Object {
        $_.Message -like "*User*"
    } | Select-Object TimeCreated, Id, LevelDisplayName, Message

$outputFile = "$PWD\Rapport_Evenements_$User.xlsx"

if ($events){
# Exporter vers un fichier Excel
    $events | Export-Excel -Path $outputFile -WorksheetName 'Événements utilisateur' -AutoSize

    Write-Host "Les événements ont été exportés dans le fichier $outputFile"
} else {
    Write-Host "Aucun événement trouvé"

}

# .\Get-UserEventsBetweenTimes.ps1 -User 'username' -StartTime "18/03/2025 12:00:00" -EndTime "18/03/2025 12:00:00"

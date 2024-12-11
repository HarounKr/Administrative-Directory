
param(
    [Parameter(Mandatory=$true)]    
    [string]$User,       # Nom de l'utilisateur
    [Parameter(Mandatory=$true)]
    [string]$StartTime,  
    [Parameter(Mandatory=$true)]
    [string]$EndTime 
)

try {
    $startTime = Get-Date $StartTime -ErrorAction Stop
    $startTime
    $endTime = Get-Date $EndTime -ErrorAction Stop
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

# .\Get-UserEventsBetweenTimes.ps1 -User 'username' -StartTime "2024-12-11 08:00:00" -EndTime "2024-12-11 18:00:00"
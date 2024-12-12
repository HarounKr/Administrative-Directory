param(
    [Parameter(Mandatory=$true)]   
    [string]$FolderPath  # Chemin complet du dossier
)

# Récupérer les logs d'accès au fichier
$events = Get-WinEvent -LogName Security -MaxEvents 50 | Where-Object {
    $_.Id -in 4416,4656,4663 -and $_.Message -like "*$FolderPath*"
} | Select-Object TimeCreated, Id, Message

$pathSplit = $FolderPath.Split("\")
$pathSplitLen = $pathSplit.Length
$folderName = $pathSplit[$pathSplitLen - 1]

$outputFile = "$PWD\Rapport_Evenements_$folderName.xlsx"


if ($events){
# Exporter vers un fichier Excel
    $events | Export-Excel -Path $outputFile -WorksheetName 'Événements' -AutoSize

    Write-Host "Les événements ont été exportés dans le fichier $outputFile"
} else {
    Write-Host "Aucun événement trouvé"

}
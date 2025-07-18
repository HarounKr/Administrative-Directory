# Scripts PowerShell d'Audit de Sécurité Windows

Ce dépôt regroupe plusieurs scripts PowerShell destinés à **l’analyse des journaux d’événements Windows** pour des besoins de sécurité, de traçabilité ou d'audit.

## Prérequis

- Windows avec PowerShell 5.1 ou plus récent.
- Accès administrateur (nécessaire pour lire certains logs de sécurité).
- Le module PowerShell `ImportExcel` doit être installé pour exporter les résultats :
  ```powershell
  Install-Module -Name ImportExcel -Force
  ```
## Résumé des Scripts

| ID | Script                              | Description                                                                 | Commande PowerShell                                                                                                          | Fichier de sortie                                |
|----|-------------------------------------|-----------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------|--------------------------------------------------|
| 1  | `Get-ConnectedUsersBetweenTimes.ps1` | Liste les utilisateurs connectés entre deux horaires.                      | `.\Get-ConnectedUsersBetweenTimes.ps1 -StartTime "18/03/2025 12:00:00" -EndTime "18/03/2025 13:00:00"`                      | `Rapport_Connexions_Utilisateurs.xlsx`           |
| 2  | `Get-FolderAccessEvents.ps1`        | Affiche les accès à un dossier spécifique (lecture, modification, etc.).   | `.\Get-FolderAccessEvents.ps1 -FolderPath "C:\MonDossier\Dossier"`                                                     | `Rapport_Evenements_Confidentiel.xlsx`           |
| 3  | `Get-SecurityIncidents.ps1`         | Récupère les incidents de sécurité (échecs, verrouillages, créations...).  | `.\Get-SecurityIncidents.ps1`                                                                                               | `Rapport_Evenements_SecurityIncidents.xlsx`      |
| 4  | `Get-UserEventsBetweenTimes.ps1`    | Liste les événements liés à un utilisateur entre deux horaires donnés.     | `.\Get-UserEventsBetweenTimes.ps1 -User "jdupont" -StartTime "18/03/2025 12:00:00" -EndTime "18/03/2025 14:00:00"`          | `Rapport_Evenements_jdupont.xlsx`                |
| 5  | `Get-UsersIP.ps1`                   | Liste les adresses IP utilisées lors des connexions réussies.              | `.\Get-UsersIP.ps1`                                                                                                         | `Rapport_Evenements_usersIP.xlsx`                |

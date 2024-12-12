3# ▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
# ██░▄▄▄█▀▄▀█▀▄▄▀█░▄▀█░▄▄▀██▄██░▄▄▀████████░▄▄▄░█░██░█▀▄▄▀█▀▄▄▀█▀▄▄▀█░▄▄▀█▄░▄███▄░▄█▄▄░▄▄
# ██░▄▄▄█░█▀█░██░█░█░█░▀▀░██░▄█░▀▀▄████████▄▄▄▀▀█░██░█░▀▀░█░▀▀░█░██░█░▀▀▄██░█████░████░██
# ██░▀▀▀██▄███▄▄██▄▄██▄██▄█▄▄▄█▄█▄▄████████░▀▀▀░██▄▄▄█░████░█████▄▄██▄█▄▄██▄████▀░▀███░██
# ▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀
# █▓▒▒░░░ Applications Interface Tests secondaires vers Airtable ░░░▒▒▓█
# 
# Lecture du fichier texte dans le dossiers D:/Master10/TestsSecondaires (ou E: pour marseille) 
# Parse du CSV en fonction d'une liste d'ent�te de colonne (En cas de modification il faudra me pr�venir) 
# Construction d'un objet array powershell avec les bonnes valeurs et champs Airtable
# Import dans Airtable 
# Suppression du fichier du dossier dans TestsSecondaires (Dans tous les cas il est dans auditSecondaires) 
# Traitement des pannes par un script Make (pour que l'�quipe digitale ai la main dessus en cas de modification des process de prod) 

# Dans les étapes d'installation il y a : 
# Copie du dossier AT.TS vers D:/Master10 
# Installation de PowerShell 7 (pour avoir l'ensemble des fonctions utilisées) 
# Installation du module import excel (install-module -name "ImportExcel") 
# Création de la tâche planifiée
# 04/12/2024 V1.0
# J. Lavault
#
# UTF-8
# Powershell 7


############################################## Variables et paramètres ##############################################
# Trace console ou fichier
# $global:TraceFichier = $False
$global:TraceFichier = $True


# Version du présent script 
$global:VersionScript = '1.0'
$global:varMainPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

# Chemin des fichiers de log
$logDirectory = "D:\Master10\AT.TS\Log\"
$cheminRecordImporte = Join-Path -Path $logDirectory -ChildPath "Record_importe.txt"
$cheminRecordNA = Join-Path -Path $logDirectory -ChildPath "Record_NA.txt"
$errorLogFile = "D:\Master10\AT.TS\Log\Error.txt"



# Chemin d'accès aux rapports des CSV
# $CheminCSVTS = "$varMainPath\RapportsBlanccoExtraits"
$CheminCSVTS = "D:\Master10\TestsSecondaires\"

# Importer le module PowerShell pour la manipulation d'Excel

Import-Module -Name 'ImportExcel'


# Paramètres de connexion à Airtable
$AirtableBaseId = "appCRLdVZ3L8ClP3b"
$OAuthAccessToken = "patGnIQbRCvBVrpK2.175df668d63b51d6a1e8df882de908a407c045d6efdc36eae2d4d38c8e195d29"
$AirtableTableRapports = "tbllozeFezAX3BDls"




###################################################### MAIN ###########################################################


function NouveauAirtableRecord {
    param (

        [Parameter(Mandatory=$true)]
        [hashtable]$Fields
    )

    # URL de l'API Airtable
    $url = "https://api.airtable.com/v0/$AirtableBaseId/$AirtableTableRapports"

    # Entête 
    $headers = @{
        "Authorization" = "Bearer $OAuthAccessToken"
        "Content-Type"  = "application/json"
    }

    # Créer le corps de la requête JSON
    $body = @{
        "fields" = $Fields
        "typecast" = $true
    } | ConvertTo-Json

    try {
        $reponserequeteAT = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body
        return $reponserequeteAT
    } catch {
        $errorMessage = "Erreur d'API Airtable : $($_.Exception.Message)"
        $errorMessage | Out-File -FilePath $errorLogFile -Append
        return "null"
    }
}

function AjoutAirtableFields {
    param (
        [Parameter(Mandatory=$true)]
        [hashtable]$Fields,

        [Parameter(Mandatory=$true)]
        [string]$FieldName,

        [Parameter(Mandatory=$true)]
        [psobject]$FieldValue
    )

    # Ajouter ou mettre à jour le champ dans le hashtable
    $Fields[$FieldName] = $FieldValue

    # Retourner le hashtable mis à jour
    return $Fields
}

function LogImportSuccess {
    param (
        [string]$txtfilename,
        [string]$recordId
    )
    $importDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $message = "$importDate - Le fichier $txtfilename a été importé avec succès. ID du record : $recordId"
    Add-Content -Path $cheminRecordImporte -Value $message
}

# Fonction pour enregistrer les messages d'erreur d'import
function LogImportError {
    param (
        [string]$txtfilename,
        [string]$errorMessage
    )
    $errorDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $message = "$errorDate - Erreur lors de l'importation du fichier $txtfilename : $errorMessage"
    Add-Content -Path $cheminRecordNA -Value $message
}

# Fonction pour lire et parser un fichier TXT
function ParseTxtFile {
    param (
        [string]$filePath
    )

    # Vérifiez que le fichier existe
    if (!(Test-Path -Path $filePath)) {
        Write-Error "Le fichier spécifié n'existe pas : $filePath"
        return $null
    }

    # Importer les données en tant qu'objet CSV
    $headers = @(
        "Numero_commande_TS", "Date_audit_ts", "Numinterne_ts", "Num_serie_ts",
        "Marque_ts", "Modele_ts", "Cpu_ts", "Ram_TS", "Taille_disque_ts", "Etat_disque_ts",
        "Etat_batterie_ts", "Etat_clavier_ts", "Test_dvd_ts", "Test_usb_ts", "Test_cam_ts",
        "Test_son_ts", "Test_microphone_ts", "Test_wifi_ts", "Divers_ts", "Test_sd_ts",
        "Info_bios_ts", "Info_licence_ts", "Plasturgie_ts", "Grade_ecran_ts", "Autonomie_batterie_ts",
        "CGrade_ts", "Observations_ts", "Pseudo_operateur_ts", "Debut_audit_ts", "Fin-audit_ts",
        "Unused", "Mise_en_veille_ts", "Test_bluetooth_ts", "Test_touchpad_ts", "Os_ts",
        "Etat_bat1_alt", "Etat_bat1_ts", "Etat_bat2_ts", "Defaut_ts"
    )

    # Charger les données avec les en-têtes
    $donnees = Import-Csv -Path $filePath -Delimiter ";" -Header $headers

    # Appliquer les transformations spécifiques
    foreach ($row in $donnees) {
        $row.Etat_bat1_ts = if (![string]::IsNullOrWhiteSpace($row.Etat_bat1_alt)) {
            $row.Etat_bat1_alt
        } else {
            $row.Etat_bat1_ts
        }
    }

    # Retourner les données transformées
    return $donnees
}
############################################################# Parcours des fichiers ###########################################################################


# Boucle de parcours des CSV de tests secondaires

# Parcourir chaque fichier TXT dans le dossier spécifié
foreach ($txtFile in Get-ChildItem -Path $CheminCSVTS -Filter *.txt) {
    try {
        # Parse le fichier TXT
        $parsedData = ParseTxtFile -filePath $txtFile.FullName

        # Créer la variable $ChampsAirtable (Hashtable)
        $ChampsAirtable = @{}
        $Import = "oui"

        # Utiliser une boucle pour ajouter les champs Airtable
        foreach ($row in $parsedData) {
            foreach ($key in $row.PSObject.Properties.Name) {
                $value = $row.$key
                if ($value -and $value -ne "null" -and $value -ne "" -and $key -ne "Unused" -and $key -ne "Etat_bat1_alt") {
                    $ChampsAirtable = AjoutAirtableFields -Fields $ChampsAirtable -FieldName $key -FieldValue $value
                }
            }
        }

        # Créer un enregistrement si $Import est "oui"
        if ($Import -ne "non") {
            $reponseAT = NouveauAirtableRecord -Fields $ChampsAirtable
        } else {
            Remove-Item -Path $txtFile.FullName -Force
        }

        # Si l'import est un succès
        if ($reponseAT.id) {
            LogImportSuccess -txtFileName $txtFile.Name -recordId $reponseAT.id
            Remove-Item -Path $txtFile.FullName -Force
        }

    } catch {
        # En cas d'erreur, journaliser l'erreur
        LogImportError -txtFileName $txtFile.Name -errorMessage $_.Exception.Message
    }
}

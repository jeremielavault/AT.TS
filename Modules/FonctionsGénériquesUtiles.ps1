# ▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄▄
# ██░▄▄▄█▀▄▀█▀▄▄▀█░▄▀█░▄▄▀██▄██░▄▄▀████████░▄▄▄░█░██░█▀▄▄▀█▀▄▄▀█▀▄▄▀█░▄▄▀█▄░▄███▄░▄█▄▄░▄▄
# ██░▄▄▄█░█▀█░██░█░█░█░▀▀░██░▄█░▀▀▄████████▄▄▄▀▀█░██░█░▀▀░█░▀▀░█░██░█░▀▀▄██░█████░████░██
# ██░▀▀▀██▄███▄▄██▄▄██▄██▄█▄▄▄█▄█▄▄████████░▀▀▀░██▄▄▄█░████░█████▄▄██▄█▄▄██▄████▀░▀███░██
# ▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀▀
# █▓▒▒░░░ Applications Interface Blancco vers Airtable ░░░▒▒▓█
# █▓▒▒░░░ Fonctions génériques utiles                  ░░░▒▒▓█
# 
# 
#
# 16/05/2024 V1.0
# P. Fayollat - J. Lavault
#
# UTF-8
# Powershell 7


############################################## Fonctions ##############################################
Function New-FileNameTimeStamped {
	# Ajouter un Timestamp au nom de fichier
	param($FileName, $Date=(Get-Date -Format 'yyyy-MM-dd HH-mm-ss'))
	
	If ([String]::IsNullOrEmpty($FileName) -eq $False) {
		If ($Date.GetType().Name -eq 'DateTime') {
			$DateFichier = $Date.ToString('yyyy-MM-dd HH-mm-ss')
		}
		Else {$DateFichier = $Date}
		$ParsedFile = New-object System.IO.FileInfo $FileName
		$TimeStampedFile = $ParsedFile.Directory.ToString().TrimEnd("\")+'\'+$ParsedFile.BaseName.ToString()+'_'+$DateFichier+$ParsedFile.Extension.ToString()
	}
	Else { $TimeStampedFile = $Null }
	
	Return $TimeStampedFile
}# endfuntion New-FileNameTimeStamped

Function Format-TimeSpan {
	# Convertir un chrono en HH:MM:SS de type chaine
    Param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [System.Diagnostics.StopWatch]$Chrono
    )
    #le sépérateur ":" est inclus
	$TimeSpan = [TimeSpan]::FromSeconds($Chrono.Elapsed.TotalSeconds)
    $Heures = $TimeSpan.Hours.ToString("00")
    $Minutes = $TimeSpan.Minutes.ToString("\:00")
    $Secondes = [Math]::Round(($TimeSpan.Seconds+($TimeSpan.Milliseconds/1000)), 0).ToString("\:00")

    Return $($Heures + $Minutes + $Secondes)
}

Function Trace-Log ([String]$Logfile, [String]$LogSaut = [char]10, [String]$LogMessage, [String]$ForeColor, [String]$BackColor, [Boolean]$FileMandatory = $False) {
	# Tracer le message $LogMessage dans un fichier de log ou à l'écran avec $LogSaut devant (CR ou CR+LF), de la couleur $Forecolor sur un fond $BackColor
    
	# Le message est précédé de son heure
	$Message = "$(Get-Date -Format 'yyyy-MM-dd HH-mm-ss') - $LogMessage"
	
	# Tester si on est en mode console et si $FileMandatory = $True l'écriture en fichier est obligatoire
	If (($null -ne $host.UI.RawUI.ForegroundColor) -And ($FileMandatory -eq $False)) { 
        If ([String]::IsNullOrEmpty($LogSaut) -eq $False) {
			Write-Host ($LogSaut)
		}
		If (([String]::IsNullOrEmpty($Forecolor)) -And ([String]::IsNullOrEmpty($BackColor))) { 
			Write-Host ($Message)
		}
		ElseIf ([String]::IsNullOrEmpty($BackColor)) {
			Write-Host -ForegroundColor $Forecolor ($Message)
		}
		ElseIf ([String]::IsNullOrEmpty($Forecolor)) {
			Write-Host -BackgroundColor $BackColor ($Message)
		}
		Else {
			Write-Host -ForegroundColor $ForeColor -BackgroundColor $BackColor ($Message)
		}
    }
	# en mode batch, on écrit la trace dans le fichier de log $LogFile
	Else { 
		Add-Content $LogFile -value ("$($LogSaut)$($Message)")
    }
	Return $Null
}

Function Move-FichierCsvVersLog ([String]$NumInterne, [String]$CheminFichierCSV, [String]$CheminFichierLOG) {
	# Déplacer le fichier portant un numéro interne déterminé dans un répertoire de log avec ajout d'un timestamp
	Get-ChildItem $CheminFichierCSV -File ([char]215+"$NumInterne"+[char]215+".csv") 
		| Move-Item –Destination {"$DestinationFichier$($_.BaseName)"+'_'+(Get-Date -Format 'yyyy-MM-dd HH-mm-ss')+'.log'}
	Return $Null
}

Function Send-NotificationSlack {
    # Envoyer une notificatins dans le fil # receptionpc d'Ecodair
    Param(
        [Parameter(Mandatory)]
        [string]$SlackMessage,

        [Parameter()]
        [string]$SlackWebhookURI
    )
    Try {
	Send-SlackMessage -Uri $SlackWebhookURI -Text $SlackMessage -ErrorAction SilentlyContinue
	}
	Catch {
		
	}
    Return $Null
}

function InstalleDernièreVersionModule {
    # Installer la dernière version d'un module depuis Internet
    param (
        # Paramètre $moduleSouhaité
        [Parameter(Mandatory)]
        [String]$moduleSouhaité
        )
    # Vérifie si le module $moduleSouhaité est installé
    $ModuleInstallés = Get-Module -Name $moduleSouhaité -ListAvailable | Sort-Object Version -Descending

    if ($ModuleInstallés) {
        # Récupère les versions installées
        $VersionInstallées = $ModuleInstallés.Version
        
        # Vérifie s'il y a une version plus récente
        $dernièreVersion = Find-Module -Name $moduleSouhaité | Select-Object -ExpandProperty Version | Sort-Object -Descending | Select-Object -First 1

        # Désinstalle les modules de version plus anciennes
        foreach ($module in $ModuleInstallés) {
            if ($module.Version -ne $dernièreVersion) {
                Uninstall-Module -Name $moduleSouhaité -RequiredVersion $module.Version -Force
                Write-Host "Version $($module.Version) du module $moduleSouhaité désinstallée."
            }
        }

        if (-not($dernièreVersion -in $VersionInstallées)) {
            Write-Host "Une version plus récente ($dernièreVersion) du module $moduleSouhaité est disponible."
            # Installe la nouvelle version
            Install-Module -Name $moduleSouhaité -Force
            Write-Host "La version $dernièreVersion du module $moduleSouhaité a été installée avec succès."
        }
        else {
            Write-Host "Le module $moduleSouhaité est déjà à la dernière version ($dernièreVersion) disponible."
        }
    }
    else {
        # Si le module n'est pas installé, installe-le
        Write-Host "Le module $moduleSouhaité n'est pas installé. Installation en cours..."
        Try {
            Install-Module -Name $moduleSouhaité -Force -ErrorAction 'SilentlyContinue'
            Write-Host "Le module $moduleSouhaité a été installé avec succès."
        }
        Catch {
            Write-Host "Le module $moduleSouhaité est introuvable."
        }
    }
    Return $Null
}
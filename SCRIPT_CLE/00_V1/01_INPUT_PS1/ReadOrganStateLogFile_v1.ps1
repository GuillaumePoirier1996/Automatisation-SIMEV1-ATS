#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	1.0
# Description 	: 	Recherche du dernier etat d'organe dans
#					les logs
#
# README 		:	lire la documentation associée
#####################################################################################################################



#INITIALISATION
Param(
		# Nom de l'organe recherche
		[Parameter(Position=0)]
		[String]
		$organ_name,
		# Horodatage de l'action
		[Parameter(Position=1)]
		$h,
		# Log d'execution du programme de recherche
		[Parameter(Position=2)]
		[String]
		$LogExeFullPath
	)

try {
        # Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
        Clear-Host
        $ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# creation du repertoire de log si non existence
		if((Test-Path -Path $LogExeFullPath) -eq $FALSE){
			New-Item -Path $LogExeFullPath -ItemType File -Force
		}

		# Enplacement du fichier
		[String] $directory = "D:\NEXTLogger\CCK"
		
		# Chaine de caractere a reconnaitre dans le nom du fichier
		[String] $filenamePattern = "Traces_MRGA_Exploit_1_"
		
		# Creation du tableau de nom des fichiers
		$filename = @()
		
		# Creation de la variable de sortie
		[String] $state_value =""
		
		# Transformation de h en format date
		$h = get-date($h)
		
		if(Test-Path -Path "$directory\$filenamePattern*") {
			'[INFO] - ' + $(Get-Date) + ' : ReadOrganStateLogFile_v1.ps1 : il existe des fichiers comportants la chaine de caractere demande' | Out-File -FilePath $LogExeFullPath -Append
			Get-Childitem "$directory\$filenamePattern*" | foreach {
				$filename += $_.name
			}
		}
		
		if ($filename.Count -eq 0) {
			# Gestion du cas ou le fichier est introuvable dans le repertoire
			'[ERROR] - ' + $(Get-Date) + ' : ReadOrganStateLogFile_v1.ps1 : impossible de trouver un fichier avec un titre composé de ' + $filenamePattern + ' dans le repertoire ' + $directory + ' pour l''organe ' + $organ_name | Out-File -FilePath $LogExeFullPath -Append
			$state_value ="4"
			exit $state_value	
		}
		
		while ($filename.Count -ne 1) {
			# Gestion du cas ou le fichier est en cours d'incrementation
			$filename = @()
			sleep 1
			Get-Childitem "$directory\$filenamePattern*" | foreach {
				$filename += $_.name
			}
		}
		
		# Creation d'un tableau contenant les lignes resultant de la recherche 
		$organ_line = @()
		
		# Creation de la variable du chemin complet
		[String] $logfile = Join-Path -Path $directory -ChildPath $filename[0]
		
		# Recherche du dernier etat de l'organe dans chaque fichier avec les horodatages correspondants
		# Pour l'instance "organ_name" on va chercher dans le fichier "logfile" (chemin complet)
		[String] $organ_line = Get-Content $logfile | Select-String -Pattern $organ_name | Select-String -Pattern "Nouvelle valeure :" | sort -Descending | Select-Object -First 1
		
		if ($organ_line -eq "" -or $organ_line -eq $NULL) {
			# Gestion du cas ou l'organe est introuvable
			'[ERROR] - ' + $(Get-Date) + ' : ReadOrganStateLogFile_v1.ps1 : impossible de trouver l''organe ' + $organ_name + ' dans les fichiers proposes' | Out-File -FilePath $LogExeFullPath -Append
			$state_value ="5"
			exit $state_value
		}
			
		# Transformation de l'horodatage de la ligne
		$h_test = $organ_line.Substring(1,19)
		$h_test = $h_test.Split(" ")
		$h_test[1] = $h_test[1].Replace("/",":")
		$h_test = $h_test[0] + " " + $h_test[1]
		
		# Test de comparaison de l'horodatage
		if ($h -ge $h_test) {
			# Separation de la ligne avec chaque caractere ":"
			$organ_line_test = $organ_line.Split(":")
			
			# L'etat prend la derniere partie
			[String] $state_value = $organ_line_test[$organ_line_test.Count -1]
			
			# Gestion du cas incoherent
			if ($state_value -eq "incohérent") {
				$state_value = "2"
			}
			
			# Gestion du cas non dynamise
			if ($state_value -eq "non dynamisé") {
				$state_value = "3"
			}
			
			# Information de l'etat de l'organe reporte dans le fichier d'execution
			'[INFO] - ' + $(Get-Date) + ' : ReadOrganStateLogFile_v1.ps1 : le dernier etat de l''organe ' + $organ_name + ' est ' + $state_value | Out-File -FilePath $LogExeFullPath -Append
			exit $state_value
		}
		
		if ($state_value -eq "" -or $state_value -eq $NULL) {
			# Gestion du cas ou l'horodatage est incoherent
			'[ERROR] - ' + $(Get-Date) + ' : ReadOrganStateLogFile_v1.ps1 : horodatage incohérent avec l''heure du changement d''etat de l''organe' + $organ_name | Out-File -FilePath $LogExeFullPath -Append
			$state_value ="6"
			exit $state_value
		}
}

catch {
		# Gestion du cas de l'erreur exeptionnelle
		'[ERROR] - ' + $(Get-Date) + ' : ReadOrganStateLogFile_v1.ps1 : Erreur exeptionnelle pour l''organe ' + $organ_name + ' : ' + $($_.exception.message) | Out-File -FilePath $LogExeFullPath -Append
		$state_value = "7"
		exit $state_value
}
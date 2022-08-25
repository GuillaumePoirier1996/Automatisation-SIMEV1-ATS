#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	1.0
# Description 	: 	Rentrer une clé dans le bandeau mode expert
#
# README 		:	lire la documentation associe
#####################################################################################################################

# Initialisation des paramètres d'entrées
PARAM (
	# Cle a rentrer dans le champ de saisie
    [Parameter(Position=0)]
    [String]
    $Key,
	# Log d'execution du programme
	[Parameter(Position=1)]
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

        # Définition de l'application à mettre en avant
        # [String] $appToOpen = [Environment]::GetEnvironmentVariable('COMPUTERNAME')
        # $appToOpen += '_1'
		[String] $appToOpen = "Menu_1"

        # Gestion d'erreur pour les caractères incorrects ou pour une clé vide
        if (($key -eq '') -or ($key -match '[^0-9 ./]')) {
	        '[ERROR] - ' + $(Get-Date) + ' : keyAction.ps1 : format de cle incorrect ou cle inexistante' | Out-File -FilePath $LogExeFullPath -Append
			$LASTEXITCODE=1
        }
        else {
			# Chargement des utilitaires
            [void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
			[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
			
			# Mise en premier plan de l'application à ouvrir
            [Microsoft.VisualBasic.Interaction]::AppActivate($appToOpen)
            
			# Action F5 dans l'application
            [System.Windows.Forms.SendKeys]::SendWait('{F5}')
            
			# Action écriture de la clée dans l'application
            [System.Windows.Forms.SendKeys]::SendWait($Key)
            
			# Action ENTER dans l'application
            [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
            
			'[INFO] - ' + $(Get-Date) + ' : keyAction.ps1 : cle ' + $Key + ' rentree' | Out-File -FilePath $LogExeFullPath -Append
			
        }
    }
catch {
		
		# Gestion des exeptions
        '[ERROR] - ' + $(Get-Date) + ' : keyAction.ps1 : ' + $($_.exception.message) | Out-File -FilePath $LogExeFullPath -Append
    }

EXIT $LASTEXITCODE

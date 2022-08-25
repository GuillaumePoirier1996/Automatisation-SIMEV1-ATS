#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	1.0
# Description 	: 	Faire une impression ecran de
#					fenetre voulue mise en avant
#
# README 		:	lire la documentation associe
#####################################################################################################################

# Définition du paramètre d'entrée : la clée
PARAM (
    # Emplacement pour le screenshot de la fenêtre
	[Parameter(Position=0)]
    [String]
    $ScreenShotPath,
	# Nom du screenshot
	[Parameter(Position=1)]
    [String]
	$ScreenShotName,
	# Application à mettre en premier plan
	[Parameter(Position=2)]
	[String]
	$appToOpen,
	# Chemin complet du log d'execution du programme
	[Parameter(Position=3)]
    [String]
    $LogExeFullPath
)


# Gestion d'erreur pour l'application introuvable
try {
        # Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# creation du repertoire de log si non existence
		if((Test-Path -Path $LogExeFullPath) -eq $FALSE){
			New-Item -Path $LogExeFullPath -ItemType File -Force
		}
		
		# creation du repertoire de screenshot si non existence
		$testSpePath = $ScreenShotName.Split("s")
		$testSpePath = $testSpePath[0]
		[String] $testSpePath = $testSpePath.Substring(0,$testSpePath.Length - 1)
		
		$stepPath = $ScreenShotName.Split("_")
		[String] $stepPath = $stepPath[6]
		
		$ScreenShotPath = $ScreenShotPath + "\" + $testSpePath + "\STEP_" + $stepPath
		
		if((Test-Path -Path $ScreenShotPath) -eq $FALSE){
			New-Item -Path $ScreenShotPath -itemType Directory -Force
		}
		
		# Chargement des utilitaires
        [void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
		[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
		

		# Mise en premier plan de l'application appToOpen
		[Microsoft.VisualBasic.Interaction]::AppActivate($appToOpen)
			
		# Action ALT + IMPR ECRAN dans l'application appToOpen
		[System.Windows.Forms.SendKeys]::SendWait("%{PRTSC}")
	
		# Enregistrement de la capture d’ecran au format bmp
		[String] $MyFile = $ScreenShotPath + "\" + $ScreenShotName + ".bmp"
		Get-Clipboard -Format Image | ForEach-Object -MemberName Save -ArgumentList $MyFile
		
		'[INFO] - ' + $(Get-Date) + ' : ScreenShot.ps1 : Screenshot ' + $ScreenShotName + ' realise avec la fenetre ' + $appToOpen + ' mise en avant enregistre dans ' + $ScreenShotPath | Out-File -FilePath $LogExeFullPath -Append
			
	}
	catch {
		EXIT
	}

EXIT
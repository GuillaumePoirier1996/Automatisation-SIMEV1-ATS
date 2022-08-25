#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	1.0
# Description 	: 	Confirmer la formation d'un itinéraire ou d'une autorisation particulier
#
# README 		:	lire la documentation associée
#####################################################################################################################

# Initialisation des paramètres d'entrées
PARAM (
	# Log d'execution du programme
	[Parameter(Position=0)]
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

        # Gestion pour la fenetre d'itinéraire particulier de l'accent
		$appToOpen = Get-Process | Where-Object {$_.MainWindowTitle -like "Confirmer la commande de l'itin*raire particulier"}  | Select -expand MainWindowTitle
		
		# Test d'existence de la fenetre
		if ($appToOpen -eq $Null){
			'[ERROR] - ' + $(Get-Date) + ' : confirmSettingUp.ps1 : fenetre absente des process en cours' | Out-File -FilePath $LogExeFullPath -Append
			$LASTEXITCODE = 1
			EXIT $LASTEXITCODE
		}
		
        # Chargement des utilitaires
        [void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
		[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        
		# Mise en premier plan de l'application à ouvrir
		[Microsoft.VisualBasic.Interaction]::AppActivate($appToOpen)

		$signature=@'
		[DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
		public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@
		
		Add-Type -AssemblyName System.Windows.Forms
		$XY = [System.Windows.Forms.Screen]::AllScreens | Where-Object {$_.Primary -eq "True"} | Select-Object WorkingArea
		$MidWidth = ($XY.WorkingArea.Width / 2) + $XY.WorkingArea.X
		$MidHeight = ($XY.WorkingArea.Height / 2) + $XY.WorkingArea.Y
		
		$SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru
		
# Action de clique souris dans l'application
		# Chargement des coordonnees
		$x = $MidWidth - 40
		$y = $MidHeight + 60
		
		# Placement du curseur aux coordonnees x et y
		[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
		# Simulation d'un clique
		$SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
		# Simulation du relachement
		$SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
		
		#Test de la fermeture de la fenetre
		$TestWindow = Get-Process | Where-Object {$_.MainWindowTitle -like "Confirmer la commande de l'itin*raire particulier"}  | Select -expand MainWindowTitle
		
		if ($TestWindow -eq $Null){
			'[INFO] - ' + $(Get-Date) + ' : confirmSettingUp.ps1 : clique ok fenetre absente des process en cours' | Out-File -FilePath $LogExeFullPath -Append
			EXIT $LASTEXITCODE
		}
		else {
			'[ERROR] - ' + $(Get-Date) + ' : confirmSettingUp.ps1 : click nok fenetre encore presente dans les process en cours' | Out-File -FilePath $LogExeFullPath -Append
			$LASTEXITCODE = 1
			EXIT $LASTEXITCODE
		}
    }
catch {
		
		# Gestion des exeptions
        '[ERROR] - ' + $(Get-Date) + ' : ' + $($_.exception.message) | Out-File -FilePath $LogExeFullPath -Append
		EXIT $LASTEXITCODE
	}



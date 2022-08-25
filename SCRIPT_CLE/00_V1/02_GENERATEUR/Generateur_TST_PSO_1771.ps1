#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# File			:	Generateur_TST_PSO_737
# Version 		: 	1.0
# Description 	: 	Generateur de la famille de test NExTEO_ATS_SFE_TST_PSO-1771 Clé 21
#
# README 		:	lire la documentation associée
#####################################################################################################################

# INITIALISATION
PARAM (
		# Chemin complet pour la BDD du poste
		[Parameter(Position=0)]
		[String]
		$BDDFullPath,
		# Chemin complet pour la FE du poste
		[Parameter(Position=1)]
		[String]
		$FEFullPath,
		# Poste concerne
		[Parameter(Position=2)]
		[String]
		$Poste,
		# Chemin racine pour le poste concerne
		[Parameter(Position=3)]
		[String]
		$PostePath,
		# Chemin racine
		[String]
		$RootPath = "D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1",
		# Chemin racine pour les utilitaires ps1
		[String]
		$ps1Path = "$RootPath\01_INPUT_PS1",
		# Chemin racine pour les utilitaires sps
		[String]
		$spsPath = "$RootPath\00_INPUT_SPS",
		# Test généré
		[String]
		$testGenere = "1771",
		# Chemin racine pour le manager de la famille de test
		[String]
		$managerPath = "$PostePath\02_SCRIPT_CLE\TST_PSO_" + $testGenere + "\00_MANAGER",
		# Chemin racine pour les scripts de la famille de test
		[String]
		$SpsDirectory = "$PostePath\02_SCRIPT_CLE\TST_PSO_" + $testGenere + "\01_SCRIPT",
		# Chemin racine pour les log de génération
		[String]
		$logPath = "$PostePath\02_SCRIPT_CLE\TST_PSO_" + $testGenere + "\02_LOG",
		# Chemin racine pour les livrables
		[String]
		$OutputPathReport = "$PostePath\03_RAPPORT_CLE\TST_PSO_" + $testGenere
)

try {
	# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
	# S'il y a une erreur le programme s'arrete
	$ErrorActionPreference = "Stop"
	
	# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
	$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
	
	# Test du contenu de BDDFullPath
	if ($BDDFullPath -eq "" -or $BDDFullPath -eq $Null) {
		# Gestion de l'erreur entree vide
		[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Generateur : l''entree BDDFullPath est vide'
		write-host $ErrorMsg
		pause
		exit
	}
	
	# Test du contenu de FEFullPath
	if ($FEFullPath -eq "" -or $FEFullPath -eq $Null) {
		# Gestion de l'erreur entree vide
		[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Generateur : l''entree FEFullPath est vide'
		write-host $ErrorMsg
		pause
		exit
	}
	
	# Test du contenu de Poste
	if ($Poste -eq "" -or $Poste -eq $Null) {
		# Gestion de l'erreur entree vide
		[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Generateur : l''entree Poste est vide'
		write-host $ErrorMsg
		pause
		exit
	}
	
	# Test du contenu de PostePath
	if ($PostePath -eq "" -or $PostePath -eq $Null) {
		# Gestion de l'erreur entree vide
		[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Generateur : l''entree PostePath est vide'
		write-host $ErrorMsg
		pause
		exit
	}

	# Fichier Temp pour la gestion des variables d'environements
	[String]$TempModuleEnv=$Env:PSModulePath
	$env:PSModulePath

	# Initialisation de la variable d'environement pour aller chercher le module / fonctions
	$Env:PSModulePath = ($Env:PSModulePath + ";" + $ps1Path)

	# Importation du module pour utiliser ses fonctions internes
	Import-Module ("$ps1Path\KeyUtility.psm1")

	get-module -ListAvailable -name KeyUtility
	
	# Création du fichier de log vide s'il n'existe pas dans le repertoire courant
	[String] $logFilename = 'Generateur_TST_PSO_' + $testGenere + "_" + $(Get-Date -format dd_MM_yyyy) +'.txt'
	[String] $LogFullPath = Join-Path -Path $logPath -ChildPath $logFilename
	if((Test-Path -Path $LogFullPath) -eq $FALSE){
		New-Item -Path $LogFullPath -ItemType File -Force
	}
	
	# Information du debut de génération
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' | Out-File -FilePath $LogFullPath -Append
	'[INFO] - ' + $(Get-Date) + ' : Generateur : DEBUT DE GENERATION' | Out-File -FilePath $LogFullPath -Append
	
	[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo] "fr-FR"
	
	# Ouvertue du fichier excel
	$global:Excel = New-Object -ComObject excel.application
	# Excel restera ferme pour l'operateur
	$global:Excel.visible = $FALSE
	# Ouverture du fichier
	$global:Workbook = $Excel.Workbooks.Open($FEFullPath,0,$TRUE)
	
	# Information de l''ouverture de la PE
	'[INFO] - ' + $(Get-Date) + ' : Generateur : la procedure d''essai ' + $PE + ' a correctement ete ouverte'| Out-File -FilePath $LogFullPath -Append
	
	# Ouverture de la feuille "PdG"
	$global:Worksheet = $Workbook.sheets.item("PdG")

	'[INFO] - ' + $(Get-Date) + ' : Generateur : ouverture de l''onglet PdG'| Out-File -FilePath $LogFullPath -Append

	# Test de la variable Poste
	[String] $posteTest = ""

	for ([int] $i = 21;$i -lt 30;$i++){
		for ([int] $j=1;$j -lt 10;$j++){
			$posteTest = $posteTest + $worksheet.Cells.Item($i,$j).Value()
		}
	}
	
	$posteTest = $posteTest.Substring($posteTest.Length - 2,2)
	
	if ($posteTest -ne $Poste) {
		# Gestion de l'erreur incohérence du poste
		'[ERROR] - ' + $(Get-Date) + ' : Generateur : le poste renseigne dans l''onglet "PdG" (' + $posteTest + ') incoherent avec le poste choisit pour la generation (' + $Poste + ')'| Out-File -FilePath $LogFullPath -Append
		pause
		exit
	}

	'[INFO] - ' + $(Get-Date) + ' : Generateur : le poste ' + $Poste + ' est le poste concerne'| Out-File -FilePath $LogFullPath -Append
	
	
	# Initialisation des images a utiliser pour les screenshots
	[String] $reminderTicket = "Tickets rappels"
	[String] $alarmTicket = "Tickets alarmes"
	[String] $buttonBox = "Test signalisation"
	
	$tcoi = @()
	$globale = @()
	
	switch ($Poste) {
		("22") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. lentes/impaires"
			$tcoi += "TCOi CHL V. lentes/impaires"
			$globale += "Image support : Globale CHL V. lentes/paires/impaires"
			$globale += "Globale CHL V. lentes/paires/impaires"
			$globale += "Image support : Globale CHL V. lentes/impaires"
			$globale += "Globale CHL V. lentes/impaires"
		}
		("24") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. paires"
			$tcoi += "TCOi CHL V. paires"
			$globale += "Image support : Globale CHL V. lentes/paires/impaires"
			$globale += "Globale CHL V. lentes/paires/impaires"
			$globale += "Image support : Globale CHL V. paires"
			$globale += "Globale CHL V. paires"
		}
		("25") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. lentes/impaires"
			$tcoi += "TCOi CHL V. lentes/impaires"
			$globale += "Image support : Globale CHL V. lentes/paires/impaires"
			$globale += "Globale CHL V. lentes/paires/impaires"
			$globale += "Image support : Globale CHL V. lentes/impaires"
			$globale += "Globale CHL V. lentes/impaires"
		}
		("26") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. paires"
			$tcoi += "TCOi CHL V. paires"
			$globale += "Image support : Globale CHL V. lentes/paires/impaires"
			$globale += "Globale CHL V. lentes/paires/impaires"
			$globale += "Image support : Globale CHL V. paires"
			$globale += "Globale CHL V. paires"
		}
		("27") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. lentes/impaires"
			$tcoi += "TCOi CHL V. lentes/impaires"
			$globale += "Image support : Globale CHL V. lentes/paires/impaires"
			$globale += "Globale CHL V. lentes/paires/impaires"
			$globale += "Image support : Globale CHL V. lentes/impaires"
			$globale += "Globale CHL V. lentes/impaires"
		}
		("75") {
			$tcoi += "Image support : TCOi central ouest"
			$tcoi += "TCOi central ouest"
			$globale += "Image support : Globale central ouest"
			$globale += "Globale central ouest"
		}
		("81") {
			$tcoi += "Image support : TCOi central ouest"
			$tcoi += "TCOi central ouest"
			$globale += "Image support : Globale central ouest"
			$globale += "Globale central ouest"
		}
		("83") {
			$tcoi += "Image support : TCOi central est"
			$tcoi += "TCOi central est"
			$globale += "Image support : Globale central est"
			$globale += "Globale central est"
		}
		("85") {
			$tcoi += "Image support : TCOi central est"
			$tcoi += "TCOi central est"
			$globale += "Image support : Globale central est"
			$globale += "Globale central est"
		}
		("87") {
			$tcoi += "Image support : TCOi central est"
			$tcoi += "TCOi central est"
			$globale += "Image support : Globale central est"
			$globale += "Globale central est"
		}
	}
	
	# Ouverture de la feuille "Introduction"
	$global:Worksheet = $Workbook.sheets.item("Introduction")

	'[INFO] - ' + $(Get-Date) + ' : Generateur : ouverture de l''onglet Introduction'| Out-File -FilePath $LogFullPath -Append
			
	# Initialisation des fichiers d'entrées
	# Création d'un objet document d'entrée
	# Format imposé : Titre / Référence / Version

	$EntryDocuments = @(
	# Commencement à la troisième ligne imposé
		For ($n=3;$worksheet.Cells.Item($n,1).Value() -ne $null;$n++){
			$Title = $worksheet.Cells.Item($n,1).Value()
			$Title+=$worksheet.Cells.Item($n,2).Value()
			$Ref = $worksheet.Cells.Item($n,3).Value()
			$Ref+=$worksheet.Cells.Item($n,4).Value()
			$Version = $worksheet.Cells.Item($n,5).Value()
			
			[PSCustomObject]@{
				"Titre"=$Title
				"Reference"=$Ref
				"Version"=$Version
			}
		}
	)

	'[INFO] - ' + $(Get-Date) + ' : Generateur : le script s''appuis sur ' + $EntryDocuments.Count + ' documents d''entree'| Out-File -FilePath $LogFullPath -Append

	# Ouverture de la feuille "Sommaire"
	$global:Worksheet = $Workbook.sheets.item("Sommaire")
	
	'[INFO] - ' + $(Get-Date) + ' : Generateur : ouverture de l''onglet Sommaire'| Out-File -FilePath $LogFullPath -Append
	
	[String] $OutputManager = "$managerPath\MANAGER_TST_PSO_P" + $Poste + "_" + $testGenere + ".sps"
	if((Test-Path -Path $OutputManager) -eq $FALSE){
		New-Item -Path $OutputManager -ItemType File -Force
	}
	
	[String] $DateGeneration = Get-Date -Format "dddd MM/dd/yyyy HH:mm K"

	# Ecriture du debut du manager
	WriteManagerStart -testGenere $testGenere -Poste $Poste -OutputManager $OutputManager -EntryDocuments $EntryDocuments -OutputPathReport $OutputPathReport -LogFullPath $LogFullPath -SpsDirectory $SpsDirectory

	# Prise en compte des conditions de formation d'itinéraire sur le poste
	SpeCondSettingUpItiAu -BDDFullPath $BDDFullPath -LogFullPath $LogFullPath
		
	# Prise en compte des conditions de destruction d'itinéraire sur le poste	
	SpeCondDestItiAu -BDDFullPath $BDDFullPath -LogFullPath $LogFullPath
	
	# Condition d'arret de la boucle for si le test générique de la cellule suivante est different
	$break1 = $False
	
	# Prise en compte dans la génération de la pièce 2D4 et suppression des itinéraires particuliers
	$break2 = $False

	# Initialisation du numero du test 
	[Int] $numTest = 0
	
	# Boucle parcourant l'ensemble du sommaire de la ligne 2 jusqu'à ce qu'un cellule soit 
	# vide avec en plus la condition suivante :
	# Si le test générique renseigné dans la cellule suivante est différent de celui testé
	# alors la boucle s'arrete
	For ([Int] $i = 2; $worksheet.Cells.Item($i,1).Value() -ne $null -and $break1 -eq $FALSE; $i++) {
		if ($worksheet.Cells.Item($i,1).Value() -eq ("NExTEO_ATS_SFE_TST_PSO-" + $TestGenere)) {
			# Compteur de test
			$numTest++
			# Transformation de l'instance
			$Instance = $worksheet.Cells.Item($i,4).Value()
			if ($Instance.Substring(0,4) -eq "ITI_") {
				$Instance = $Instance.Split("_")
				$Instance = $Instance[1]
			}
			if ($Instance.Substring(0,2) -eq "Au" -or $Instance.Substring(0,2) -eq "AU") {
				$Instance = $Instance.Split("uU")
				$Instance = $Instance[1]
			}
			# Comparaison a un itineraire de la 2D4
			For ([Int] $j = 0; $j -lt $global:ItiDestSpeTerms.Count; $j++) {
				if ($Instance -eq $global:ItiDestSpeTerms[$j]."Itineraire") {
				# Condition d'ecriture dans le mananger et du script specifique
				$break2 = $True
				# Arret de la boucle for
				break
				}
			}
			
			# Si l'instance ne concerne pas un itinéraire de la 2D4
			if ($break2 -eq $False) {
				# Identifiant du test spécifique
				[String] $IdTestSpe = $worksheet.Cells.Item($i,2).Value()
				# Nom du test spécifique racourci pour les onglet et les noms des scripts
				[String] $TestSpe = $IdTestSpe.Replace("-","_")
				[String] $TestSpe = $TestSpe.Substring(15,$TestSpe.Length - 15)
				# Test générique
				[String] $IdTestGen = $worksheet.Cells.Item($i,1).Value()
				# Nom du test
				[String] $TestName = $worksheet.Cells.Item($i,3).Value()
				# Instance complete
				$InstanceCell = $worksheet.Cells.Item($i,4).Value()
				# Elements associés
				$Elements = $worksheet.Cells.Item($i,5).Value()
				# Moyen d'essai
				[String] $TestMedium = $worksheet.Cells.Item($i,6).Value()
				#Version de l'ATS
				[String] $ATSVersion = $worksheet.Cells.Item($i,7).Value()
				
				'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' | Out-File -FilePath $LogFullPath -Append
				'[INFO] - ' + $(Get-Date) + ' : Generateur : DEBUT DE GENERATION DU TEST ' + $TestSpe| Out-File -FilePath $LogFullPath -Append
				
				# Traitement sur les éléments
				$Elements = $Elements.split(";")
				
				# Initialisation des objets contenant les informations necessaire pour instancier correctement le test
				$Protection = @()
				$TsProtection = @()
				$TsList = @()
				$OrganList = @()
				$trackDeviceList = @()
				
				[Bool] $TestCOIN = $FALSE
				
				for ([Int] $p=0; $p -lt $Elements.Count; $p++) {
					switch -Wildcard ($Elements[$p]) {
						("*FNIT*") {
							# Ts pour le controle optique FNIT PRS
							$FNITFNAU = $Elements[$p]
							if ($FNITFNAU.Substring(0,8) -eq "FNIT_PRS"){
								$FNITFNAU = $FNITFNAU.Split("_")
								$FNITFNAU = $FNITFNAU[$FNITFNAU.Length - 1]
								[String] $TsFNITFNAU = "S_CCK_MFCO_SIG_" + $FNITFNAU + "_FNIT_PRS_ASPECT_CTRL"
								$TsList += $TsFNITFNAU
							}
							# Ts pour le controle optique FNIT
							else {
								$FNITFNAU = $FNITFNAU.Split("_")
								$FNITFNAU = $FNITFNAU[$FNITFNAU.Length - 1]
								[String] $TsFNITFNAU = "S_CCK_MFCO_SIG_" + $FNITFNAU + "_FNIT_ASPECT_CTRL"
								$TsList += $TsFNITFNAU
							}
						}
						("*FNAU*") {
							# Ts pour le controle optique FNAU2
							$FNITFNAU = $Elements[$p]
							if ($FNITFNAU.Substring(0,6) -eq "FNAU2_"){
								$FNITFNAU = $FNITFNAU.Split("_")
								$FNITFNAU = $FNITFNAU[$FNITFNAU.Length - 1]
								[String] $TsFNITFNAU = "S_CCK_MFCO_FNAU_" + $FNITFNAU + "_FNAU2_ASPECT_CTRL"
								$TsList += $TsFNITFNAU
							}
							# Ts pour le controle optique FNAU
							else {
								$FNITFNAU = $FNITFNAU.Split("_")
								$FNITFNAU = $FNITFNAU[$FNITFNAU.Length - 1]
								[String] $TsFNITFNAU = "S_CCK_MFCO_FNAU_" + $FNITFNAU + "_FNAU_ASPECT_CTRL"
								$TsList += $TsFNITFNAU
							}
						}
						("*TK_CO*") {
							# Organe COIN
							if ($Elements[$p].Substring(0,7) -eq "TK_COIN"){
								$TestCOIN = $TRUE
								$COIN = $Elements[$p].split("(")
								$COIN = $COIN[$COIN.length - 1]
								$COIN = $COIN.replace(")","")
								$OrganList += $COIN
							}
							# Organe CO
							else {
								$CO = $Elements[$p].split("(")
								$CO = $CO[$CO.length - 1]
								$CO = $CO.replace(")","")
								$OrganList += $CO
							}
						}
						("*TK_EIT*") {
							# Organe EIT
							$EitEau = $Elements[$p].split("(")
							$EitEau = $EitEau[$EitEau.Length - 1]
							$EitEau = $EitEau.replace(")","")
							$OrganList += $EitEau
						}
						("*TK_EAU*") {
							# Organe EAU
							$EitEau = $Elements[$p].split("(")
							$EitEau = $EitEau[$EitEau.Length - 1]
							$EitEau = $EitEau.replace(")","")
							$OrganList += $EitEau
						}
						("*TC_EIT*") {
							# TC EIT
							$TcEitEau = $Elements[$p].split("(")
							$TcEitEau = $TcEitEau[$TcEitEau.Length - 1]
							$TcEitEau = $TcEitEau.replace(")","")
						}
						("*TC_EAU*") {
							# TC EAU
							$TcEitEau = $Elements[$p].split("(")
							$TcEitEau = $TcEitEau[$TcEitEau.Length - 1]
							$TcEitEau = $TcEitEau.replace(")","")
						}
					}
				}

				# Variable necessaire pour le FNIT / FNAU
				[String] $TsItiAuActiveState = "JAUNE"
				[String] $TsItiAuActiveValue = "2"
				[String] $TsItiAuInactiveState = "GRIS_DE_PAYNE"
				[String] $TsItiAuInactiveValue = "1"
				
				# Variable necessaire pour l'alarme 3502
				[String] $AlarmId = "3502"
				[String] $AlarmLabel = "[\""$TcEitEau\""]"
				
				# Creation du script specifique
				[String] $OutputScript = "$SpsDirectory\TST_PSO_P" + $Poste + "_" + $testGenere + "_" + $numTest + ".sps"
				if((Test-Path -Path $OutputScript) -eq $FALSE) {
					New-Item -Path $OutputScript -ItemType File -Force
				}
				
				[String] $DateGeneration = Get-Date -Format "dddd MM/dd/yyyy HH:mm K"
				
				'[INFO] - ' + $(Get-Date) + ' : Generateur : DEBUT DE GENERATION DU SCRIPT ' + $IdTestSpe | Out-File -FilePath $LogFullPath -Append
				
				# Ecriture du debut du script
				WriteScriptStart -IdTestGen $IdTestGen -IdTestSpe $IdTestSpe -TestSpe $TestSpe -numTest $numTest -InstanceCell $InstanceCell -TestName $TestName -TestMedium $TestMedium -ATSVersion $ATSVersion -Poste $Poste -OutputScript $OutputScript -EntryDocuments $EntryDocuments -LogFullPath $LogFullPath
				
				# Ecriture du test d'existence des Ts
				WriteExistTest -TsList $TsList -OrganList $OrganList -trackDeviceList $trackDeviceList -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de l'état initial du script
				WriteScriptInit -IdTestSpe $IdTestSpe -Poste $Poste -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				[Int] $numStep = 1

### EIT EAU INACTIF ET BLOQUE				

				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "COMMANDE DE L'ORGANE $EitEau A INACTIF ET BLOQUAGE" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Commande de l'organe eit / eau a inactif
				WriteOrganeManualAction -organName $EitEau -command "CONTROL_STATE_TO_INACTIVE" -display "COMMANDE DE L'ORGANE $EitEau A INACTIF" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Bloquage de eit / eau
				WriteOrganeManualAction -organName $EitEau -command "BLOCK" -display "BLOQUAGE DE L'ORGANE $EitEau" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Prise de l'horodatage de l'action
				(	("		date ho = ho.now();"),
					("		string ho1 = ho.asString();"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type log CCK
				WriteLogVerifBeforeAction -organName $EitEau -organState "0" -h "ho1" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $buttonBox
				[String] $WindowName = $buttonBox
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				# Ecriture du resultat et de la fin du step
				[String] $exp = "(Verif$EitEau)" 
				WriteStepEnd -exp $exp -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				$numStep++
### CLE 21				
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 21" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ecriture de la cle
				WriteKeySettingUp -Key "21" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath -poste $Poste
				
				# Ajustement du delai et prise de l'horodatage de la clé
				(	("		date hKey = hKey.now();"),
					("		string hKey1 = hKey.asString();"),
					("		delay(30);"),
					("")
				) >> $OutputScript
				
				# Reserve tickets rappels
				(	("		println(""RESERVE TICKETS RAPPELS"");"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type Ts
				WriteTSVerif -ts $TsFNITFNAU -tsState $TsItiAuInactiveState -tsValue $TsItiAuInactiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de la vérification type alarme avec argument
				WriteAlarmVerif1 -h "hKey1" -idAlarm $AlarmId -Arg $AlarmLabel -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				For ([int] $s = 0; $s -lt $globale.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_Globale_" + $s
					[String] $WindowName = $globale[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
			
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $alarmTicket
				[String] $WindowName = $alarmTicket
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $reminderTicket
				[String] $WindowName = $reminderTicket
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp "((keyTest == 0) && (confirmSetUpTest == 0) && Verif$TsFNITFNAU && testAlarmVerif1)" -numStep $numStep -comment "Test realise automatiquement - Reserve tickets rappels" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				$numStep++
				
### EIT EAU DEBLOQUE ET CLE 24

				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "DEBLOQUAGE DE L'ORGANE $EitEau ET CLE 24" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Debloquage de eit / eau
				WriteOrganeManualAction -organName $EitEau -command "UNBLOCK" -display "DEBLOQUAGE DE L'ORGANE $EitEau" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de la cle
				WriteKeyDest -Key "24" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ajustement du delai
				(	("		delay(5);"),
					("")
				) >> $OutputScript
				
				# Reserve tickets rappels
				(	("		println(""RESERVE TICKETS RAPPELS"");"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type Ts
				WriteTSVerif -ts $TsFNITFNAU -tsState $TsItiAuInactiveState -tsValue $TsItiAuInactiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				For ([int] $s = 0; $s -lt $globale.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_Globale_" + $s
					[String] $WindowName = $globale[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $reminderTicket
				[String] $WindowName = $reminderTicket
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp "((keyTest == 0) && Verif$TsFNITFNAU)" -numStep $numStep -comment "Test realise automatiquement - Reserve tickets rappels" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				$numStep++

### CLE 21				
				
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 21" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ecriture de la cle
				WriteKeySettingUp -Key "21" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath -poste $Poste
				
				# Ajustement du delai et prise de l'horodatage de la clé
				(	("		date hKey = hKey.now();"),
					("		string hKey1 = hKey.asString();"),
					("		delay(12);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type Ts
				WriteTSVerif -ts $TsFNITFNAU -tsState $TsItiAuActiveState -tsValue $TsItiAuActiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de la vérification type log CCK
				WriteLogVerif -organName $CO -organState "1" -h "hKey1" -OutputScript $OutputScript -LogFullPath $LogFullPath

				[String] $exp = "((keyTest == 0) && (confirmSetUpTest == 0) && Verif$TsFNITFNAU && Verif$CO"
				
				if ($TestCOIN -eq $TRUE) {
					# Ecriture de la vérification type log CCK
					WriteLogVerif -organName $COIN -organState "1" -h "hKey1" -OutputScript $OutputScript -LogFullPath $LogFullPath
					$exp = $exp + " && Verif$COIN"
				}
				
				$exp = $exp + ")"
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				For ([int] $s = 0; $s -lt $globale.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_Globale_" + $s
					[String] $WindowName = $globale[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
			
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $buttonBox
				[String] $WindowName = $buttonBox
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp $exp -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath

### RETOUR A L'ETAT INITIAL
### CLE 24
					
				# Ecriture de la cle
				WriteKeyDest -Key "24" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath
			
				# Ajustement du delai
				(	("		delay(5);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la fin du script
				WriteScriptEnd -TestSpe $TestSpe -Poste $Poste -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				'[INFO] - ' + $(Get-Date) + ' : Generateur : FIN DE GENERATION DU TEST ' + $TestSpe | Out-File -FilePath $LogFullPath -Append
				'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------' | Out-File -FilePath $LogFullPath -Append
				
				# Ecriture de l'appel du test spécifique dans le manager
				WriteManagerScript -numTest $numTest -TestSpe $TestSpe -InstanceCell $InstanceCell -OutputManager $OutputManager -LogFullPath $LogFullPath
			}
			
			# Si l'id de test générique suivant est différent de celui testé
			if($worksheet.Cells.Item($i,1).Value() -ne $worksheet.Cells.Item($i + 1,1).Value()) {
				$break1 = $True
			}
			
		} 
	}

	# Ecriture de la fin du manager
	WriteManagerEnd -testGenere $testGenere -Poste $Poste -OutputManager $OutputManager -LogFullPath $LogFullPath
	
	
	# Creation du script specifique
	[String] $OutputInitAllEnvVar = "$managerPath\INIT_ALL_ENV_VAR.sps"
	if((Test-Path -Path $OutputInitAllEnvVar) -eq $FALSE) {
		New-Item -Path $OutputInitAllEnvVar -ItemType File -Force
	}
	
	# Ecriture du fichier d'initialisation des variables d'environnement
	WriteInitAllEnvVar -SpsDirectory $SpsDirectory -OutputInitAllEnvVar $OutputInitAllEnvVar -LogFullPath $LogFullPath
	
	# Information de la fin de génération
	'[INFO] - ' + $(Get-Date) + ' : Generateur : FIN DE GENERATION' | Out-File -FilePath $LogFullPath -Append
	'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'| Out-File -FilePath $LogFullPath -Append
	
	# Fermeture d'excel
	$global:Excel.Workbooks.Close()
	$global:Excel.Quit()
	
	# Remise à l'état initial des variables d'environements
	$Env:PSModulePath=$TempModuleEnv

}

Catch {
	
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Generateur : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
		# Fermeture d'excel
		$Excel.Workbooks.Close()
		$Excel.Quit()
		# Remise à l'état initial des variables d'environements
		$Env:PSModulePath=$TempModuleEnv
}
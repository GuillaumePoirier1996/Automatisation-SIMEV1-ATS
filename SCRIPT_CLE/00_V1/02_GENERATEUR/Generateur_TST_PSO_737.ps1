#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# File			:	Generateur_TST_PSO_737
# Version 		: 	1.0
# Description 	: 	Generateur de la famille de test NExTEO_ATS_SFE_TST_PSO-737 Clé 34 : Décondamnation d'un 
#					itinéraire/autorisation
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
		$testGenere = "737",
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
	[String] $buttonBox = "Test signalisation"
	[String] $alarmTicket = "Tickets alarmes"
	$tcoi = @()
	$protecIm = @()
	
	switch ($Poste) {
		("22") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. lentes/impaires"
			$tcoi += "TCOi CHL V. lentes/impaires"
			$protecIm += "Image support : Protections CHL V. lentes/paires/impaires"
			$protecIm += "Protections CHL V. lentes/paires/impaires"
			$protecIm += "Image support : Protections CHL V. lentes/impaires"
			$protecIm += "Protections CHL V. lentes/impaires"
		}
		("24") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. paires"
			$tcoi += "TCOi CHL V. paires"
			$protecIm += "Image support : Protections CHL V. lentes/paires/impaires"
			$protecIm += "Protections CHL V. lentes/paires/impaires"
			$protecIm += "Image support : Protections CHL V. paires"
			$protecIm += "Protections CHL V. paires"
		}
		("25") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. lentes/impaires"
			$tcoi += "TCOi CHL V. lentes/impaires"
			$protecIm += "Image support : Protections CHL V. lentes/paires/impaires"
			$protecIm += "Protections CHL V. lentes/paires/impaires"
			$protecIm += "Image support : Protections CHL V. lentes/impaires"
			$protecIm += "Protections CHL V. lentes/impaires"
		}
		("26") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. paires"
			$tcoi += "TCOi CHL V. paires"
			$protecIm += "Image support : Protections CHL V. lentes/paires/impaires"
			$protecIm += "Protections CHL V. lentes/paires/impaires"
			$protecIm += "Image support : Protections CHL V. paires"
			$protecIm += "Protections CHL V. paires"
		}
		("27") {
			$tcoi += "Image support : TCOi CHL V. lentes/paires/impaires"
			$tcoi += "TCOi CHL V. lentes/paires/impaires"
			$tcoi += "Image support : TCOi CHL V. lentes/impaires"
			$tcoi += "TCOi CHL V. lentes/impaires"
			$protecIm += "Image support : Protections CHL V. lentes/paires/impaires"
			$protecIm += "Protections CHL V. lentes/paires/impaires"
			$protecIm += "Image support : Protections CHL V. lentes/impaires"
			$protecIm += "Protections CHL V. lentes/impaires"
		}
		("75") {
			$tcoi += "Image support : TCOi central ouest"
			$tcoi += "TCOi central ouest"
			$protecIm += "Image support : Protections central ouest"
			$protecIm += "Protections central ouest"
		}
		("81") {
			$tcoi += "Image support : TCOi central ouest"
			$tcoi += "TCOi central ouest"
			$protecIm += "Image support : Protections central ouest"
			$protecIm += "Protections central ouest"
		}
		("83") {
			$tcoi += "Image support : TCOi central est"
			$tcoi += "TCOi central est"
			$protecIm += "Image support : Protections central est"
			$protecIm += "Protections central est"
		}
		("85") {
			$tcoi += "Image support : TCOi central est"
			$tcoi += "TCOi central est"
			$protecIm += "Image support : Protections central est"
			$protecIm += "Protections central est"
		}
		("87") {
			$tcoi += "Image support : TCOi central est"
			$tcoi += "TCOi central est"
			$protecIm += "Image support : Protections central est"
			$protecIm += "Protections central est"
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
				$TsList = @()
				$OrganList = @()
				$trackDeviceList = @()
				
				[Bool] $TestRV = $FALSE
				[Bool] $TestCOIN = $FALSE
				[Bool] $TestZA = $FALSE
				[Bool] $TestFC = $FALSE
				[Bool] $TestZAP = $FALSE
				
				for ([Int] $p=3; $p -lt $Elements.Count; $p++) {
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
						("*TK_RV*") {
							# Organe RV
							$TestRV = $TRUE
							$RV = $Elements[$p].split("(")
							$RV = $RV[$RV.Length - 1]
							$RV = $RV.replace(")","")
							$OrganList += $RV
							# Ts pour le TK RV
							$z = $RV.Split("V")
							$z = $z[$z.length - 1]
							[String] $TsTkRv = "S_CCK_MRGA_" + $Poste + "_CDV_z" + $z + "_LIBRE"
							$TsList += $TsTkRv
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
						("*z*") {
							# Zone amont à commander
							if (($Elements[$p].Substring(0,3) -ne "ZEP") -and ($Elements[$p].Substring(0,3) -ne "ZAP")) {
								$TestZA = $TRUE
								$ZA = "CDV_" + $Elements[$p]
								$trackDeviceList += $ZA
							}
						}
						("*FC_*") {
							# Booleen test FC
							$TestFC = $TRUE
							# Ts pour le controle optique FC
							$FC = $Elements[$p]
							$FC = $FC.Split("_")
							$FC = $FC[$FC.Length - 1]
							[String] $TsFC ="S_CCK_MFCO_SIG_" + $FC + "_FC_ASPECT_CTRL"
							$TsList += $TsFC
						}
						("*ZAP_*") {
							# Booleen test ZAP
							$TestZAP = $TRUE
							# Ts pour le controle optique ZAP
							$ZAP = $Elements[$p]
							$ZAP = $ZAP.Split("_")
							$ZAP = $ZAP[$ZAP.Length - 1]
							[String] $TsZAP = "S_CCK_MFCO_ZAP_" + $ZAP + "_ZAP_ASPECT_CTRL"
							$TsList += $TsZAP
							# Ts pour le TK ZAP
							[String] $TsTkZap = "S_CCK_MRGA_" + $Poste + "_ZAP_" + $ZAP +"_LIBRE"
							$TsList += $TsTkZap
							
						}
					}
				}
						
				# Signal de depart en position 2
				$SigD = $Elements[2]

				# Protection type ZEP obligatoirement avec le test 737 et plus de problème de decondamnation du COIN uniquement car groupement selectionné pour tout couvrir
				$Protection = $Elements[$Elements.Length - 1]
				$Protection = $Protection.Split("_")
				$Protection = $Protection[$Protection.Length - 1]
				# Ts pour le controle optique ZEP
				[String] $TsProtection = "S_CCK_MFCO_ZEP_" + $Protection + "_ZEP_LABEL_ASPECT_CTRL"
				$TsList += "S_CCK_MFCO_ZEP_" + $Protection + "_ZEP_LABEL_ASPECT_CTRL"

				# Variable necessaire pour la FC
				if ($TestFC -eq $TRUE) {
					[String] $FcActiveState = "ROUGE"
					[String] $FcActiveValue = "2"
				}
				
				# Variable necessaire pour la ZAP
				if ($TestZAP -eq $TRUE) {
					[String] $ZapInactiveState = "ROUGE_CLIGNOTANT"
					[String] $ZapInactiveValue = "2"
					
					[String] $TkZapInactiveState = "NON"
					[String] $TkZapInactiveValue = "0"
				}
				
				# Variable necessaire pour le RV
				if ($TestRV -eq $TRUE) {
					[String] $TkRvInactiveState = "NON"
					[String] $TkRvInactiveValue = "0"
				}

				# Variable necessaire pour le FNIT / FNAU
				[String] $TsItiAuActiveState = "JAUNE"
				[String] $TsItiAuActiveValue = "2"
				[String] $TsItiAuInactiveState = "GRIS_DE_PAYNE"
				[String] $TsItiAuInactiveValue = "1"
				
				# Variable necessaire pour l'avertissement 3148
				[String] $AvertLabel = "\""C_CCK_MCMD_SIG_" + $SigD + "_DECONDAMNER\"""
				[String] $AvertId = "3148"
				
				# Variable necessaire pour la Protection
				[String] $ProtActiveState = "ROUGE"
				[String] $ProtActiveValue = "2"
				
				# Variable necessaire pour l'alarme 3546
				[String] $AlarmId = "3546"
				
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
### CLE 25 SI FC				
				if ($TestFC -eq $TRUE) {
					# Ecriture de la presentation du step
					WriteStepPres -numStep $numStep -displayStep "CLE 25" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ecriture de la cle
					[String] $Key25 = "25 " + $SigD
					WriteKey -Key $Key25 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ajustement du delai
					(	("		delay(3);"),
						("")
					) >> $OutputScript
					
					# Ecriture de la vérification type Ts
					WriteTSVerif -ts $TsFC -tsState $FcActiveState -tsValue $FcActiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ecriture des screenshots
					For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
						[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
						[String] $WindowName = $tcoi[$s]
						WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
					}
					
					# Ecriture du resultat et de la fin du step
					WriteStepEnd -exp "((keyTest1 == 0) && Verif$TsFC)" -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					$numStep++
				}
### CLE 21				
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 21" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ecriture de la cle
				WriteKeySettingUp -Key "21" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath -poste $Poste
				
				# Ajustement du delai
				(	("		delay(5);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type Ts
				WriteTSVerif -ts $TsFNITFNAU -tsState $TsItiAuActiveState -tsValue $TsItiAuActiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp "((keyTest == 0) && (confirmSetUpTest == 0) && Verif$TsFNITFNAU)" -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				$numStep++
### CLE 34				
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 34" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de la cle
				[String] $Key34 = "34 " + $SigD
				WriteKey -Key $Key34 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ajustement du delai et prise de l'horodatage de la clé
				(	("		date hKey = hKey.now();"),
					("		string hKey1 = hKey.asString();"),
					("		delay(1.5);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type avertissement avec TC
				WriteAvertVerif1 -h "hKey1" -idAvert $AvertId -idCommand $AvertLabel -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ecriture des screenshots
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $alarmTicket
				[String] $WindowName = $alarmTicket
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp "((keyTest1 == 0) && testAvertVerif1)" -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				$numStep++
### CLE 24			
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 24" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ecriture de la cle
				WriteKeyDest -Key "24" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ajustement du delai
				(	("		delay(5);"),
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
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp "((keyTest == 0) && Verif$TsFNITFNAU)" -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				$numStep++
				
### CLE 31				
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 31" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de la cle
				[String] $Key31 = "31 " + $Protection
				WriteKey -Key $Key31 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ajustement du delai
				(	("		delay(2);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type Ts
				WriteTSVerif -ts $TsProtection -tsState $ProtActiveState -tsValue $ProtActiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				For ([int] $s = 0; $s -lt $protecIm.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_Protections_" + $s
					[String] $WindowName = $protecIm[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				# Ecriture du resultat et de la fin du step
				[String] $exp = "((keyTest1 == 0) && Verif" + $TsProtection + ")"
				WriteStepEnd -exp $exp -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				$numStep++
				
### DOUBLE CLE 21				
					
				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 21" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ecriture de la cle
				WriteKeySettingUp -Key "21" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath -poste $Poste
				
				# Ajustement du delai et création d'un bouléen pour l'execution de la première clé
				(	("		bool keyExe = ((keyTest == 0) && (confirmSetUpTest == 0));"),
					("		delay(10.5);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la deuxième cle
				WriteKeySettingUp -Key "21" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath -poste $Poste
				
				# Ajustement du delai et création d'un bouléen pour l'execution de la première clé
				(	("		bool keyExe = (keyExe && (keyTest == 0) && (confirmSetUpTest == 0));"),
					("		delay(5);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type Ts
				WriteTSVerif -ts $TsFNITFNAU -tsState $TsItiAuActiveState -tsValue $TsItiAuActiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp "(keyExe && Verif$TsFNITFNAU)" -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				$numStep++
				
### CLE 34 SI ZA
				if ($TestZA -eq $TRUE) {
					# Ecriture de la presentation du step
					WriteStepPres -numStep $numStep -displayStep "CLE 34" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ecriture de la cle
					[String] $Key34 = "34 " + $SigD
					WriteKey -Key $Key34 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
						
					# Ajustement du delai et prise de l'horodatage de la clé
					(	("		date hKey = hKey.now();"),
						("		string hKey1 = hKey.asString();"),
						("		delay(3);"),
						("")
					) >> $OutputScript
					
					# Ecriture de la vérification type alarme sans argument
					WriteAlarmVerif2 -h "hKey1" -idAlarm $AlarmId -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ecriture des screenshots
					For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
						[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
						[String] $WindowName = $tcoi[$s]
						WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
					}
					
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $alarmTicket
					[String] $WindowName = $alarmTicket
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
					
					# Ecriture du resultat et de la fin du step
					WriteStepEnd -exp "((keyTest1 == 0) && testAlarmVerif2)" -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					$numStep++
					
### OCCUPER LA ZONE AMONT SI ZA

					# Ecriture de la presentation du step
					WriteStepPres -numStep $numStep -displayStep "OCCUPER LA ZONE AMONT" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# OCCUPATION DE LA ZONE AMONT PAR ACTION A PIED D'OEUVRE
					WriteOrganeManualAction -organName $ZA -command "INCIDENT_FORCE_OCCUPIED" -display "DERANGEMENT : FORCAGE DU CDV $ZA A L ETAT OCCUPE" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ajustement du delai prise de l'horodatage de l'action a pied d'oeuvre et gestion de l'apparition de train
					(	("		date ho = ho.now();"),
						("		string ho1 = ho.asString();"),
						(""),
						("		delay(3);"),
						(""),
						("		use TRAIN_MOTION_EMULATOR;"),
						("		int numberOfTrains = TRAIN_MOTION_EMULATOR.getNumberOfTrains();"),
						("		if (numberOfTrains > 0) {"),
						("			variant train = TRAIN_MOTION_EMULATOR.getTrainByNumber(numberOfTrains);"),
						("			TRAIN_MOTION_EMULATOR.deleteTrain(train.getId());"),
						("			delay(2);"),
						("		}"),
						("")
					) >> $OutputScript
					
### SI ZAP

					if ($TestZAP -eq $TRUE) {
						# Ecriture de la vérification type Ts
						WriteTSVerif -ts $TsZAP -tsState $ZapInactiveState -tsValue $ZapInactiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
						[String] $exp = "(Verif$TsZAP"
						
						# Ecriture de la vérification type Ts
						WriteTSVerif -ts $TsTkZap -tsState $TkZapInactiveState -tsValue $TkZapInactiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
						$exp = $exp + " && Verif$TsTkZap"
### SI RV
						if ($TestRV -eq $TRUE) {
							# Ecriture de la vérification type Ts
							WriteTSVerif -ts $TsTkRv -tsState $TkRvInactiveState -tsValue $TkRvInactiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
							$exp = $exp + " && Verif$TsTkRv" 
						}
						
						$exp = $exp + ")"
					}
					
### SANS ZAP

					else {
						
### SI RV

						if ($TestRV -eq $TRUE) {
							# Ecriture de la vérification type Ts
							WriteTSVerif -ts $TsTkRv -tsState $TkRvInactiveState -tsValue $TkRvInactiveValue -OutputScript $OutputScript -LogFullPath $LogFullPath
							[String] $exp = "Verif$TsTkRv" 
						}
						
### SANS RV
						else {
							(	("	// ZONE AMONT SANS RV NI ZAP DONC IMPOSSIBLE DE VERIFIER"),
								("		println(""ZONE AMONT SANS RV NI ZAP DONC IMPOSSIBLE DE VERIFIER"");")
							) >> $OutputScript
							[String] $exp = "false"
						}
					}	
					
					# Ecriture des screenshots
					For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
						[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s
						[String] $WindowName = $tcoi[$s]
						WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
					}
					
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $alarmTicket
					[String] $WindowName = $alarmTicket
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
					
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $buttonBox
					[String] $WindowName = $buttonBox
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
					
					# Ecriture du resultat et de la fin du step
					WriteStepEnd -exp $exp -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					$numStep++
				}
				
### CLE 34

				# Ecriture de la presentation du step
				WriteStepPres -numStep $numStep -displayStep "CLE 34" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ecriture de la cle
				[String] $Key34 = "34 " + $SigD
				WriteKey -Key $Key34 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
				# Ajustement du delai et prise de l'horodatage de la clé
				(	("		date hKey = hKey.now();"),
					("		string hKey1 = hKey.asString();"),
					("		delay(12);"),
					("")
				) >> $OutputScript
				
				# Ecriture de la vérification type log CCK
				WriteLogVerif -organName $CO -organState "1" -h "hKey1" -OutputScript $OutputScript -LogFullPath $LogFullPath
				[String] $exp = "(Verif$CO"
				
### SI COIN
				if ($TestCOIN -eq $TRUE) {
					# Ecriture de la vérification type log CCK
					WriteLogVerif -organName $COIN -organState "1" -h "hKey1" -OutputScript $OutputScript -LogFullPath $LogFullPath
					$exp = $exp + " && Verif$COIN"
				}
				
				$exp = $exp +")"
				
				# Ecriture des screenshots
				For ([int] $s = 0; $s -lt $tcoi.Count; $s++) {
					[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_TCOi_" + $s 
					[String] $WindowName = $tcoi[$s]
					WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				}
				
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $alarmTicket
				[String] $WindowName = $alarmTicket
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				[String] $ScreenShotName = $TestSpe + "_step_" + $numStep + "_" + $buttonBox
				[String] $WindowName = $buttonBox
				WriteScreenShot -ScreenShotName $ScreenShotName -OutputScript $OutputScript -WindowName $WindowName -LogFullPath $LogFullPath
				
				# Ecriture du resultat et de la fin du step
				WriteStepEnd -exp $exp -numStep $numStep -comment "Test realise automatiquement" -numTest $numTest -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				$numStep++
				
### RETOUR A L'ETAT INITIAL
### SI ZA

				if ($TestZA -eq $TRUE) {
					WriteOrganeManualAction -organName $ZA -command "RAISE_INCIDENT" -display "DERANGEMENT LEVE POUR LE CDV $ZA" -OutputScript $OutputScript -LogFullPath $LogFullPath
					
					# Ajustement du delai
					(	("		delay(3);"),
						("")
					) >> $OutputScript
				}
				
### CLE 24
				# Ecriture de la cle
				WriteKeyDest -Key "24" -Instance $Instance -OutputScript $OutputScript -LogFullPath $LogFullPath
			
				# Ajustement du delai
				(	("		delay(5);"),
					("")
				) >> $OutputScript

### CLE 32 ET 33

				[String] $Key32 = "32 " + $Protection
				WriteKey -Key $Key32 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ajustement du delai
				(	("		delay(2);"),
					("")
				) >> $OutputScript
				
				[String] $Key33 = "33 " + $Protection
				WriteKey -Key $Key33 -i "1" -OutputScript $OutputScript -LogFullPath $LogFullPath
				
				# Ajustement du delai
				(	("		delay(2);"),
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
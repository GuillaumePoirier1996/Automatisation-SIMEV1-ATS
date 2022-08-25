#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	1.0
# Description 	: 	Module des utilitaires pour les clés
#
# README 		:	lire la documentation associée
#####################################################################################################################

# Decommenter test des dates 
# Fonction permetant d'enregistrer les conditions de formations des Itineraires / Autorisations sur un poste en particulier
Function SpeCondSettingUpItiAu {
	# Initialisation des paramètres d'entrées
	PARAM (
		# Chemin complet pour la BDD du poste
		[Parameter(Position=0)]
		[String]
		$BDDFullPath,
		# Chemin complet pour le log d'execution du programme
		[Parameter(Position=1)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de BDDFullPath
		if ($BDDFullPath -eq "" -or $BDDFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : l''entree BDDFullPath est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo] "fr-FR"

		# Ouvertue de la BDD Poste
		$Excel = New-Object -ComObject excel.application
		# Excel restera ferme pour l'operateur
		$Excel.visible = $FALSE

		# Ouverture du fichier en lecture seule
		$Workbook = $Excel.Workbooks.Open($BDDFullPath,0,$TRUE)
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : la BDD ' + $BDDFullPath	+ ' a ete ouverte' | Out-File -FilePath $LogFullPath -Append
			
		# Test de conformite de la BDD
		$Worksheet = $Workbook.sheets.item("PdG")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : ouverture de l''onglet PdG' | Out-File -FilePath $LogFullPath -Append
		
		# Test du titre
		[String] $titleTest = ""

		for ([int] $i = 21;$i -lt 30;$i++){
			for ([int] $j=1;$j -lt 10;$j++){
				$titleTest += $worksheet.Cells.Item($i,$j).Value()
			}
		}
		
		$titleBDD = $BDDFullPath.Split("\")
		[String] $titleBDD = $TitleBDD[$TitleBDD.Length - 1]
		
		if ($titleTest.Substring(0,15) -ne "BASE DE DONNEES" -or $titleTest.Substring(16,5) -ne "POSTE") {
			# Gestion de l'erreur de conformité du titre par rapport à "BASE DE DONNEES POSTE"
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : Titre du document non conforme par rapport a ''BASE DE DONNEES POSTE''' | Out-File -FilePath $LogFullPath -Append
			# Fermeture d'excel
			$Excel.Workbooks.Close()
			$Excel.Quit()
			exit
		}
		
		# Test de la cohérence du poste
		$testPoste = $titleTest.Split(" ")
		[String] $testPoste = $testPoste[$testPoste.Length - 1]
		
		$PosteBDD = $titleBDD.Split(" ")
		$PosteBDD = $PosteBDD[2]
		$PosteBDD = $PosteBDD.Split("_")
		[String] $PosteBDD = $PosteBDD[0]
		
		if ($PosteBDD -ne $testPoste) {
			# Gestion de l'erreur d'incohérence du poste
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : Poste incoherent entre le nom du document et celui renseigne sur la page de garde' | Out-File -FilePath $LogFullPath -Append
			# Fermeture d'excel
			$Excel.Workbooks.Close()
			$Excel.Quit()
			exit
		}
		
		# Test du type de document
		[String] $docType = ""
		
		for ([int] $j=1;$j -lt 10;$j++){
			$docType += $worksheet.Cells.Item(17,$j).Value()
		}

		if ($docType -ne "DONNEES POSTE") {
			# Gestion de l'erreur d'incohérence du poste
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : Type de document non conforme' | Out-File -FilePath $LogFullPath -Append
			# Fermeture d'excel
			$Excel.Workbooks.Close()
			$Excel.Quit()
			exit
		}
		
		# Test de l'edition revision et prise de la date sur l'onglet PdG
		# Traitement sur le nom du fichier
		$testVersion = $titleTest.Split(" ")
		[String] $testVersion = $testPoste[$testPoste.Length - 1]
		
		$VersionBDD = $titleBDD.Split(" ")
		$VersionBDD = $VersionBDD[2]
		$VersionBDD = $VersionBDD.Split("_")
		[String] $VersionBDD = $VersionBDD[1]
		
		$edBDD = $VersionBDD.Split(".")
		$edBDD = $edBDD[0]
		[Int] $edBDD = $edBDD.Substring(1)
		
		$revBDD = $VersionBDD.Split(".")
		[Int] $revBDD = $revBDD[1]
		
		# Traitement su l'onglet PdG
		for ([int] $j=5;$j -lt 10;$j++){
			$versionPdG += $worksheet.Cells.Item(4,$j).Value()
		}
		
		$versionPdG = $versionPdG.Split(": ")
		[String] $versionPdG = $versionPdG[$versionPdG.Length - 1]
		
		$edPdG = $versionPdG.Split("/")
		[Int] $edPdG = $edPdG[0]
		
		$revPdG = $versionPdG.Split("/")
		[Int] $revPdG = $revPdG[1]
		
		# Date 1 de l'onglet PdG
		[String] $datePdG1 = ""
		
		for ([int] $i = 36;$i -lt 39;$i++){
			for ([int] $j = 8;$j -lt 10;$j++){
				$datePdG1 += $worksheet.Cells.Item($i,$j).Value()
			}
		}
		
		# Date 2 de l'onglet PdG
		for ([int] $j=5;$j -lt 10;$j++){
			$datePdG2 += $worksheet.Cells.Item(6,$j).Value()
		}
		
		$datePdG2 = $datePdG2.Split(": ")
		[String] $datePdG2 = $datePdG2[$datePdG2.Length - 1]
		
		# Traitement sur l'onglet Introduction et prise de la date sur l'onglet Introduction
		$Worksheet = $Workbook.sheets.item("Introduction")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : ouverture de l''onglet Introduction' | Out-File -FilePath $LogFullPath -Append
		
		[Bool] $Break = $False
		for ([Int] $i = 2; $worksheet.Cells.Item($i,1).Value() -ne $null -and -$Break -eq $False; $i++) {
			
			if ($worksheet.Cells.Item($i + 1,1).Value() -eq $null) {
				$i +=13
			}
			
			if ($worksheet.Cells.Item($i,1).Value() -eq "SUIVI D'EVOLUTIONS") {
				for ([Int] $j = $i + 2; $worksheet.Cells.Item($j,1).Value() -ne $null -and $Break -eq $False; $j++) {
					# Version complete
					[String] $versionIntro = $worksheet.Cells.Item($j,1).Value()
					# Date de la version
					[String] $dateIntro = $worksheet.Cells.Item($j,3).Value()
					if ($worksheet.Cells.Item($j + 1,1).Value() -eq $null) {
						$Break = $True
					}
				}
			}
		}
				
		$edIntro = $versionPdG.Split("/")
		[Int] $edIntro = $edIntro[0]
		
		$revIntro = $versionPdG.Split("/")
		[Int] $revIntro = $revIntro[1]
		
		if ($edIntro -ne $edPdG -or $edIntro -ne $edBDD -or $edPdG -ne $edBDD) {
			# Gestion de l'erreur sur les difference d'édition du document
			'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : Les indices d''edition sur la page de garde, sur l''introduction et contenu dans le nom du document sont differents' | Out-File -FilePath $LogFullPath -Append
		}
		
		if ($revIntro -ne $revPdG -or $revIntro -ne $revBDD -or $revPdG -ne $revBDD) {
			# Gestion de l'erreur sur les difference d'édition du document
			'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : Les indices de revision sur la page de garde, sur l''introduction et contenu dans le nom du document sont differents' | Out-File -FilePath $LogFullPath -Append
		}

		# Test de la date
		$datePdG1 = [Datetime]::ParseExact($datePdG1, 'dd/MM/yyyy', $null)
		$datePdG2 = [Datetime]::ParseExact($datePdG2, 'dd/MM/yyyy', $null)
		$dateBDD = [Datetime] (Get-ItemProperty -Path $BDDFullPath -Name LastWriteTime).LastWriteTime
		
		if ($datePdG1 -ne $datePdG2 -or $datePdG1 -ne $dateIntro -or $datePdG1 -ne $dateBDD -or $datePdG2 -ne $dateIntro -or $datePdG2 -ne $dateBDD -or $dateIntro -ne $dateBDD) {
			# Gestion de l'erreur sur les difference de dates entre la page de garde, l'introduction et la date du document
			'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : Les dates ne sont pas toutes egales entre la derniere edition du document, la page de garde et l''introduction' | Out-File -FilePath $LogFullPath -Append
		}
	
		# Ouverture de l'onglet PT 2A pour les conditions de formations des Itinéraires
		$Worksheet = $Workbook.sheets.item("PT2A")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : ouverture de l''onglet PT2A' | Out-File -FilePath $LogFullPath -Append
		# Initialisation des conditions de formations specifiques
		# Création d'un objet avec les conditions de formation
		# Format imposé : Itineraire / Condition de Formation / Renvoi

		$global:ItiSettingUpTerms = @(
			# Commencement à la ligne 4
			For ([Int] $i=4; $worksheet.Cells.Item($i,2).Value() -ne $NULL; $i++) {
				# Itinéraire dans la colonne 2 de l'onglet
				$Iti = $worksheet.Cells.Item($i,2).Value()
				# Conditions de formations de l'itinéraire dans la colonne 13 de l'onglet
				$Terms = $worksheet.Cells.Item($i,13).Value()
				# Renvoi pour l'itinéraire dans la colonne 1 de l'onglet
				# Utile en particulier pour le renvoi 10 en particulier
				$Reference = $worksheet.Cells.Item($i,1).Value()
				
				[PSCustomObject]@{
					"Itineraire"=$Iti
					"Condition de Formation"=$Terms
					"Renvoi"=$Reference
				}
				
			}
		)
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : conditions particulieres de formation d''itineraires enregistrees' | Out-File -FilePath $LogFullPath -Append
		
		# Ouverture de l'onglet PT 2B pour les conditions de formations des Autorisations
		$Worksheet = $Workbook.sheets.item("PT2B")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : ouverture de l''onglet PT2B' | Out-File -FilePath $LogFullPath -Append
		# Initialisation des conditions de formations specifiques
		# Création d'un objet avec les conditions de formation
		# Format imposé : Autorisation / Condition de Formation / Renvoi

		$global:AuSettingUpTerms = @(
			# Commencement à la ligne 4
			For ([Int] $i=4; $worksheet.Cells.Item($i,2).Value() -ne $NULL; $i++) {
				# Autorisation dans la colonne 2 de l'onglet
				$Au = $worksheet.Cells.Item($i,2).Value()
				$AuTr = $Au.Split("AuU ")
				$AuTr = $AuTr[$AuTr.Length - 1]
				
				# Conditions de formations de l'autorisation dans la colonne 10 de l'onglet
				$Terms = $worksheet.Cells.Item($i,10).Value()
				
				# Renvoi pour l'autorisation dans la colonne 1 de l'onglet
				# Utile en particulier pour le renvoi 10 en particulier
				$Reference = $worksheet.Cells.Item($i,1).Value()
				
				[PSCustomObject]@{
					"Autorisation"=$AuTr
					"Condition de Formation"=$Terms
					"Renvoi"=$Reference
				}
				
			}
		)
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : conditions particulieres de formation d''autorisations enregistrees' | Out-File -FilePath $LogFullPath -Append
		
		# Fermeture d'excel
		$Excel.Workbooks.Close()
		$Excel.Quit()
	}
	
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondSettingUpItiAu : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    	# Fermeture d'excel
		$Excel.Workbooks.Close()
		$Excel.Quit()
	}
}

# Decommenter test des dates 
# Fonction permetant d'enregistrer les conditions de destruction des Itineraires / Autorisations sur un poste en particulier
Function SpeCondDestItiAu {
	# Initialisation des paramètres d'entrées
	PARAM (
		# Chemin complet pour la BDD du poste
		[Parameter(Position=0)]
		[String]
		$BDDFullPath,
		# Chemin complet pour le log d'execution du programme
		[Parameter(Position=1)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de BDDFullPath
		if ($BDDFullPath -eq "" -or $BDDFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : l''entree BDDFullPath est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo] "fr-FR"

		# Ouvertue de la BDD Poste
		$Excel = New-Object -ComObject excel.application
		# Excel restera ferme pour l'operateur
		$Excel.visible = $FALSE

		# Ouverture du fichier en lecture seule
		$Workbook = $Excel.Workbooks.Open($BDDFullPath,0,$TRUE)
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : la BDD ' + $BDDFullPath	+ ' a ete ouverte' | Out-File -FilePath $LogFullPath -Append

		# Test de conformite de la BDD
		$Worksheet = $Workbook.sheets.item("PdG")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : ouverture de l''onglet PdG' | Out-File -FilePath $LogFullPath -Append
		
		# Test du titre
		[String] $titleTest = ""

		for ([int] $i = 21;$i -lt 30;$i++){
			for ([int] $j=1;$j -lt 10;$j++){
				$titleTest += $worksheet.Cells.Item($i,$j).Value()
			}
		}
		
		$titleBDD = $BDDFullPath.Split("\")
		[String] $titleBDD = $TitleBDD[$TitleBDD.Length - 1]
		
		if ($titleTest.Substring(0,15) -ne "BASE DE DONNEES" -or $titleTest.Substring(16,5) -ne "POSTE") {
			# Gestion de l'erreur de conformité du titre par rapport à "BASE DE DONNEES POSTE"
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Titre du document non conforme par rapport a ''BASE DE DONNEES POSTE''' | Out-File -FilePath $LogFullPath -Append
			# Fermeture d'excel
			$Excel.Workbooks.Close()
			$Excel.Quit()
			exit
		}
		
		# Test de la cohérence du poste
		$testPoste = $titleTest.Split(" ")
		[String] $testPoste = $testPoste[$testPoste.Length - 1]
		
		$PosteBDD = $titleBDD.Split(" ")
		$PosteBDD = $PosteBDD[2]
		$PosteBDD = $PosteBDD.Split("_")
		[String] $PosteBDD = $PosteBDD[0]
		
		if ($PosteBDD -ne $testPoste) {
			# Gestion de l'erreur d'incohérence du poste
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Poste incoherent entre le nom du document et celui renseigne sur la page de garde' | Out-File -FilePath $LogFullPath -Append
			# Fermeture d'excel
			$Excel.Workbooks.Close()
			$Excel.Quit()
			exit
		}
		
		# Test du type de document
		[String] $docType = ""
		
		for ([int] $j=1;$j -lt 10;$j++){
			$docType += $worksheet.Cells.Item(17,$j).Value()
		}

		if ($docType -ne "DONNEES POSTE") {
			# Gestion de l'erreur d'incohérence du poste
			'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Type de document non conforme' | Out-File -FilePath $LogFullPath -Append
			# Fermeture d'excel
			$Excel.Workbooks.Close()
			$Excel.Quit()
			exit
		}
		
		# Test de l'edition revision et prise de la date sur l'onglet PdG
		# Traitement sur le nom du fichier
		$testVersion = $titleTest.Split(" ")
		[String] $testVersion = $testPoste[$testPoste.Length - 1]
		
		$VersionBDD = $titleBDD.Split(" ")
		$VersionBDD = $VersionBDD[2]
		$VersionBDD = $VersionBDD.Split("_")
		[String] $VersionBDD = $VersionBDD[1]
		
		$edBDD = $VersionBDD.Split(".")
		$edBDD = $edBDD[0]
		[Int] $edBDD = $edBDD.Substring(1)
		
		$revBDD = $VersionBDD.Split(".")
		[Int] $revBDD = $revBDD[1]
		
		# Traitement su l'onglet PdG
		for ([int] $j=5;$j -lt 10;$j++){
			$versionPdG += $worksheet.Cells.Item(4,$j).Value()
		}
		
		$versionPdG = $versionPdG.Split(": ")
		[String] $versionPdG = $versionPdG[$versionPdG.Length - 1]
		
		$edPdG = $versionPdG.Split("/")
		[Int] $edPdG = $edPdG[0]
		
		$revPdG = $versionPdG.Split("/")
		[Int] $revPdG = $revPdG[1]
		
		# Date 1 de l'onglet PdG
		[String] $datePdG1 = ""
		
		for ([int] $i = 36;$i -lt 39;$i++){
			for ([int] $j = 8;$j -lt 10;$j++){
				$datePdG1 += $worksheet.Cells.Item($i,$j).Value()
			}
		}
		
		# Date 2 de l'onglet PdG
		for ([int] $j=5;$j -lt 10;$j++){
			$datePdG2 += $worksheet.Cells.Item(6,$j).Value()
		}
		
		$datePdG2 = $datePdG2.Split(": ")
		[String] $datePdG2 = $datePdG2[$datePdG2.Length - 1]
		
		# Traitement sur l'onglet Introduction et prise de la date sur l'onglet Introduction
		$Worksheet = $Workbook.sheets.item("Introduction")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : ouverture de l''onglet Introduction' | Out-File -FilePath $LogFullPath -Append
		
		[Bool] $Break = $False
		for ([Int] $i = 2; $worksheet.Cells.Item($i,1).Value() -ne $null -and -$Break -eq $False; $i++) {
			
			if ($worksheet.Cells.Item($i + 1,1).Value() -eq $null) {
				$i +=13
			}
			
			if ($worksheet.Cells.Item($i,1).Value() -eq "SUIVI D'EVOLUTIONS") {
				for ([Int] $j = $i + 2; $worksheet.Cells.Item($j,1).Value() -ne $null -and $Break -eq $False; $j++) {
					# Version complete
					[String] $versionIntro = $worksheet.Cells.Item($j,1).Value()
					# Date de la version
					[String] $dateIntro = $worksheet.Cells.Item($j,3).Value()
					if ($worksheet.Cells.Item($j + 1,1).Value() -eq $null) {
						$Break = $True
					}
				}
			}
		}
				
		$edIntro = $versionPdG.Split("/")
		[Int] $edIntro = $edIntro[0]
		
		$revIntro = $versionPdG.Split("/")
		[Int] $revIntro = $revIntro[1]
		
		if ($edIntro -ne $edPdG -or $edIntro -ne $edBDD -or $edPdG -ne $edBDD) {
			# Gestion de l'erreur sur les difference d'édition du document
			'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Les indices d''edition sur la page de garde, sur l''introduction et contenu dans le nom du document sont differents' | Out-File -FilePath $LogFullPath -Append
		}
		
		if ($revIntro -ne $revPdG -or $revIntro -ne $revBDD -or $revPdG -ne $revBDD) {
			# Gestion de l'erreur sur les difference d'édition du document
			'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Les indices de revision sur la page de garde, sur l''introduction et contenu dans le nom du document sont differents' | Out-File -FilePath $LogFullPath -Append
		}

		# Test de la date
		$datePdG1 = [Datetime]::ParseExact($datePdG1, 'dd/MM/yyyy', $null)
		$datePdG2 = [Datetime]::ParseExact($datePdG2, 'dd/MM/yyyy', $null)
		$dateBDD = [Datetime] (Get-ItemProperty -Path $BDDFullPath -Name LastWriteTime).LastWriteTime
		
		if ($datePdG1 -ne $datePdG2 -or $datePdG1 -ne $dateIntro -or $datePdG1 -ne $dateBDD -or $datePdG2 -ne $dateIntro -or $datePdG2 -ne $dateBDD -or $dateIntro -ne $dateBDD) {
			# Gestion de l'erreur sur les difference de dates entre la page de garde, l'introduction et la date du document
			'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Les dates ne sont pas toutes egales entre la derniere edition du document, la page de garde et l''introduction' | Out-File -FilePath $LogFullPath -Append
		}

		# Ouverture de l'onglet PT 2D2 pour les Enclenchements de Parcours des Itinéraires
		$Worksheet = $Workbook.sheets.item("PT2D2")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : ouverture de l''onglet PT2D2' | Out-File -FilePath $LogFullPath -Append
		# Initialisation des EPA
		# Création d'un objet avec les EPA
		# Format imposé : Itineraire / Temporisation / Signal de depart

		$global:ItiDestEPA = @(
			# Commencement à la ligne 4
			For ([Int] $i=4; $worksheet.Cells.Item($i,4).Value() -ne $NULL; $i++) {
				# Itinéraire dans la colonne 4 de l'onglet
				$Iti = $worksheet.Cells.Item($i,4).Value()
				# Separation de tous les itineraires avec le separateur virgule
				$Iti = $Iti.Split(",")

				# Temporisation a la destruction dans la colonne 5 de l'onglet (en minute)
				[String] $Temp = $worksheet.Cells.Item($i,5).Value()

				# Signal de depart de l'itineraire dans la colonne 3 de l'onglet
				[String] $SigD = $worksheet.Cells.Item($i,3).Value()
				$SigDCv = $SigD.Split("Cv")
				$SigDCv = $SigDCv[$SigDCv.Length - 1]
				
				For ([int] $n = 0; $n -lt $Iti.Count; $n++) {
					# Gestion de l'espace au debut de l'itineraire (genere a cause des separateurs)
					# Si le premier terme de l'itineraire est un espace
					if ($Iti[$n].Substring(0,1) -eq " "){
						# On prend uniquement la suite des caracteres (sans le permier qui est un espace)
						$Iti[$n] = $Iti[$n].Substring(1,$Iti[$n].Length - 1)
					}
					[PSCustomObject]@{
						"Itineraire"=$Iti[$n]
						"Temporisation (min)"=$Temp
						"Signal de depart"=$SigDCv
					}
				}
			}
		)
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : EPA enregistres' | Out-File -FilePath $LogFullPath -Append
		
		# Ouverture de l'onglet PT 2D3 pour les Destructions Manuelles Temporisees des Itinéraires
		$Worksheet = $Workbook.sheets.item("PT2D3")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : ouverture de l''onglet PT2D3' | Out-File -FilePath $LogFullPath -Append
		# Initialisation des DMT
		# Création d'un objet avec les conditions de formation
		# Format imposé : Itineraire / Temporisation / Signal de depart

		$global:ItiDestDMT = @(
			# Commencement à la ligne 4
			For ([Int] $i=4; $worksheet.Cells.Item($i,5).Value() -ne $NULL; $i++) {
				# Itinéraire dans la colonne 5 de l'onglet
				$Iti = $worksheet.Cells.Item($i,5).Value()
				# Separation de tous les itineraires avec le separateur virgule
				$Iti = $Iti.split(",")
				# Temporisation a la destruction dans la colonne 6 de l'onglet (en minute)
				[String] $Temp = $worksheet.Cells.Item($i,6).Value()
				
				For ([int] $n = 0; $n -lt $Iti.Count; $n++) {
					# Gestion de l'espace au debut de l'itineraire (genere a cause des separateurs)
					# Si le premier terme de l'itineraire est un espace
					if ($Iti[$n].Substring(0,1) -eq " "){
						# On prend uniquement la suite des caracteres (sans le permier qui est un espace)
						$Iti[$n] = $Iti[$n].Substring(1,$Iti[$n].Length - 1)
					}
					[PSCustomObject]@{
						"Itineraire"=$Iti[$n]
						"Temporisation (min)"=$Temp
					}
				}
			}
		)
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : DMT enregistrees' | Out-File -FilePath $LogFullPath -Append
		
		# Ouverture de l'onglet PT 2D4 pour les Itineraires avec des conditions de destruction particulieres 
		$Worksheet = $Workbook.sheets.item("PT2D4")
		'[INFO] - ' + $(Get-Date) + ' ouverture de l''onglet PT2D4' | Out-File -FilePath $LogFullPath -Append
		# Initialisation des conditions particulieres de destruction
		# Création d'un objet avec les conditions particulieres
		# Format imposé : Itineraire - Autorisation / Transit non en action / Autres Conditions / Observations

		$global:ItiDestSpeTerms = @(
			# Commencement à la ligne 4
			For ([Int] $i=4; $worksheet.Cells.Item($i,6).Value() -ne $NULL; $i++) {
				# Itinéraire dans la colonne 6 de l'onglet
				$Iti = $worksheet.Cells.Item($i,6).Value()
				# Separation de tous les itineraires avec le separateur virgule
				$Iti = $Iti.split(",")
				# Transit non en action dans la colonne 4 de l'onglet
				[String] $Transit = $worksheet.Cells.Item($i,4).Value()
				# Autres coditions dans la colonne 5 de l'onglet
				[String] $Terms = $worksheet.Cells.Item($i,5).Value()
				# Obesrvations dans la colonne 7 de l'onglet
				[String] $Obs = $worksheet.Cells.Item($i,7).Value()
				
				For ([int] $n = 0; $n -lt $Iti.Count; $n++) {
					# Gestion de l'espace au debut de l'itineraire (genere a cause des separateurs)
					# Si le premier terme de l'itineraire est un espace
					if ($Iti[$n].Substring(0,1) -eq " "){
						# On prend uniquement la suite des caracteres (sans le permier qui est un espace)
						$Iti[$n] = $Iti[$n].Substring(1,$Iti[$n].Length - 1)
					}
					[PSCustomObject]@{
						"Itineraire"=$Iti[$n]
						"Transit non en action"=$Transit
						"Autres Conditions"=$Terms
						"Observations"=$Obs
					}
				}
			}
		)
		
		'[WARNING] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Conditions Particulieres enregistrees des itineraire font l''objet de principes particuliers' | Out-File -FilePath $LogFullPath -Append
		
		# Ouverture de l'onglet PT 2B pour les Autorisation avec des conditions de destruction particulieres 
		$Worksheet = $Workbook.sheets.item("PT2B")
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : ouverture de l''onglet PT2B' | Out-File -FilePath $LogFullPath -Append
		# Initialisation des conditions particulieres de destruction
		# Création d'un objet avec les conditions particulieres
		# Format imposé : Autorisation / Zones Liberees / Autres Conditions

		$global:AuDest = @(
			# Commencement à la ligne 4
			For ([Int] $i=4; $worksheet.Cells.Item($i,2).Value() -ne $NULL; $i++) {
				# Autorisation dans la colonne 19 de l'onglet
				[String] $Au = $worksheet.Cells.Item($i,2).Value()
				$AuTr = $Au.Split("AuU ")
				$AuTr = $AuTr[$AuTr.Length - 1]
				
				[PSCustomObject]@{
					"Autorisation"=$AuTr
				}
			}
		)
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : Destructions d''autorisations enregistrees' | Out-File -FilePath $LogFullPath -Append
		
		# Fermeture d'excel
		$Excel.Workbooks.Close()
		$Excel.Quit()
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction SpeCondDestItiAu : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    }
}

# Fonction permettant d'ecrire les action pour la formation d'itinéraire ou la prise d'une autorisation
# Utilisée pour les clé 21 22 et 37.1
Function WriteKeySettingUp {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Clé à rentrer dans le bandeau mode expert
		[Parameter(Position=0)]
		[String]
		$Key,
		# Instance pour la construction de la clé
		# Utilisée pour reconnaitre l'itinéraire ou l'autorisation en question
		[Parameter(Position=1)]
		[String]
		$Instance,
		# Chemin complet du Script de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath,
		# Poste concerne
		[Parameter(Position=4)]
		[String]
		$Poste
	)

	try {
		
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
			
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de Key
		if ($Key -eq $null -or $Key -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : l''entree Key est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de Instance
		if ($Instance -eq $null -or $Instance -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : l''entree Instance est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de Poste
		if ($Poste -eq $null -or $Poste -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : l''entree Poste est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test de condition de formation des itinéraires
		# Boucle parcourant l'objet comportant les conditions
		for ([int] $i = 0; $i -lt $ItiSettingUpTerms.Count; $i++){
			# Test de correspondance des instances
			if ($ItiSettingUpTerms[$i]."Itineraire" -eq $Instance) {
				# S'il y a des conditions de formations
				if($null -ne $ItiSettingUpTerms[$i]."Condition de Formation"){
					# Traitement des condition de formation avec le separateur ","
					$AllTerms = $ItiSettingUpTerms[$i]."Condition de Formation"
					$AllTerms = $AllTerms.split(",")
					# Boucle parcourant l'ensemble des conditions de formation pour une instance
					For ([int] $n = 0; $n -lt $AllTerms.Count; $n++) {
						# S'il y a une condition LITAG
						switch -Wildcard ($AllTerms[$n]){
							("*LITAG*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("[")
								$Organe = $Organe[$Organe.Length - 1]
								$Organe = $Organe.Substring(5, $Organe.Length - 5)
								$Organe = $Organe.Replace("]","")
								$Organe = $Organe.Replace(" ","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "KPGEP" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION LITAG"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition LITAG ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition DIT 14
							("*DIT*(14)*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("DIT")
								$Organe = $Organe[0]
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.replace("[","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "PCO" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION DIT 14"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition DIT(14) ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition AUE reçue
							("*Au.E*re*ue*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split(" ")
								if ($Organe.Count -eq 4){
									$Organe = $Organe[2]
								}
								if ($Organe.Count -eq 3){
									$Organe = $Organe[1]
								}
								$Organe = $Organe.ToUpper()
								$Organe = "AUE_" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""ACTION_GIVE_ENTRY_AUTHORIZATION"",""ACTION SUR L'AUTORISATION " + $Organe + " POUR LA CONDITION AUE RECUE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition AUE RECUE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition AUS reçue
							("*Au.S*re*ue*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split(" ")
								if ($Organe.Count -eq 4){
									$Organe = $Organe[2]
								}
								if ($Organe.Count -eq 3){
									$Organe = $Organe[1]
								}
								$Organe = $Organe.ToUpper()
								$Organe = "AUS_" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""ACTION_GIVE_EXIT_AUTHORIZATION"",""ACTION SUR L'AUTORISATION " + $Organe + " POUR LA CONDITION AUS RECUE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition AUS RECUE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition transit libéré
							("*Tr*lib*r*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("Tr")
								$Organe = $Organe[$Organe.Length - 2]
								if ($Organe.Contains("P")){
									$Organe = $Organe.split("P")
									$Organe = $Organe[0]
									$Organe = $Organe.replace(" ","")
									$Organe = $Poste + "TRP" + $Organe
								}
								if ($Organe.Contains("I")){
									$Organe = $Organe.split("I")
									$Organe = $Organe[0]
									$Organe = $Organe.replace(" ","")
									$Organe = $Poste + "TRI" + $Organe
								}
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION TRANSIT LIBERE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition TRANSIT LIBERE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition Cv non origine
							("*Cv*non origine*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("Cv")
								$Organe = $Organe[$Organe.Length - 1]
								$Organe = $Organe.split(" ")
								$Organe = $Organe[0]
								$Organe = $Poste + "EIT" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION CV NON ORIGINE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition CV NON ORIGINE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition Cv non intermédiaire
							("*Cv*non interm*diaire*DIT*(15)*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("Cv")
								$Organe = $Organe[$Organe.Length - 1]
								$Organe = $Organe.split(" ")
								$Organe = $Organe[0]
								$Organe = $Poste + "EIT" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION CV NON INTERMEDIAIRE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition CV NON INTERMEDIAIRE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition FNITO
							("*FNITO*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("FNITO")
								$Organe = $Organe[$Organe.Length -1]
								$Organe = $Organe.replace("]","")
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "AXAUE" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION FNITO"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition FNITO ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition FNITD
							("*FNITD*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("FNITD")
								$Organe = $Organe[$Organe.Length -1]
								$Organe = $Organe.replace("]","")
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "AXAUS" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION FNITD"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition FNITD ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
						}
					}
				}
				
				# traitement sur l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Traitement du cas pour les itinéraires particuliers
				if($ItiSettingUpTerms[$i]."Renvoi" -eq "(10)"){
					$Instance = $Instance + " 0"
					[Bool] $ItiAuPart = $TRUE
				}
				else {
					[Bool] $ItiAuPart = $FALSE
				}
				
				break
			}
		}

		# Test de condition de prise des autorisations
		# Boucle parcourant l'objet comportant les conditions
		 for ([int] $i = 0; $i -lt $AuSettingUpTerms.Count; $i++){
			# Test de correspondance des instances
			if ($AuSettingUpTerms[$i]."Autorisation" -eq $Instance) {
<#
# AUCUNES CONDITIONS SPECIFIQUES SUR LE DOMAINE ATS
				# S'il y a des conditions de formations
				if($null -ne $AuSettingUpTerms[$i]."Condition de Formation"){
					# Traitement des condition de formation avec le separateur ","
					$AllTerms = $AuSettingUpTerms[$i]."Condition de Formation"
					$AllTerms = $AllTerms.split(",")
					# Boucle parcourant l'ensemble des conditions de formation pour une instance
					For ([int] $n = 0; $n -lt $AllTerms.Count; $n++) {
						switch -Wildcard ($AllTerms[$n]){
							# S'il y a une condition LITAG
							("*LITAG*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("LITAG")
								$Organe = $Organe[$Organe.Length -1]
								$Organe = $Organe.replace("]","")
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "KPGEP" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION LITAG"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition LITAG ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition DIT 14
							("*DIT*(14)*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("DIT")
								$Organe = $Organe[0]
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.replace("[","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "PCO" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION DIT 14"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition DIT(14) ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition DIT 15
							("*DIT*(15)*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("DIT")
								$Organe = $Organe[0]
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.replace("[","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "PIN" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION DIT 15"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition DIT(15) ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition AUE reçue
							("*Au.E*re*ue*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split(" ")
								if ($Organe.Count -eq 4){
									$Organe = $Organe[2]
								}
								if ($Organe.Count -eq 3){
									$Organe = $Organe[1]
								}
								$Organe = $Organe.ToUpper()
								$Organe = "AUE_" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""ACTION_GIVE_ENTRY_AUTHORIZATION"",""ACTION SUR L'AUTORISATION " + $Organe + " POUR LA CONDITION AUE RECUE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition AUE RECUE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition AUS reçue
							("*Au.S*re*ue*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split(" ")
								if ($Organe.Count -eq 4){
									$Organe = $Organe[2]
								}
								if ($Organe.Count -eq 3){
									$Organe = $Organe[1]
								}
								$Organe = $Organe.ToUpper()
								$Organe = "AUS_" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""ACTION_GIVE_EXIT_AUTHORIZATION"",""ACTION SUR L'AUTORISATION " + $Organe + " POUR LA CONDITION AUS RECUE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition AUS RECUE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition transit libéré
							("*Tr*lib*r*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("Tr")
								$Organe = $Organe[$Organe.Length - 2]
								if ($Organe.Contains("P")){
									$Organe = $Organe.split("P")
									$Organe = $Organe[0]
									$Organe = $Organe.replace(" ","")
									$Organe = $Poste + "TRP" + $Organe
								}
								if ($Organe.Contains("I")){
									$Organe = $Organe.split("I")
									$Organe = $Organe[0]
									$Organe = $Organe.replace(" ","")
									$Organe = $Poste + "TRI" + $Organe
								}
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION TRANSIT LIBERE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition TRANSIT LIBERE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition Cv non origine
							("*Cv*non origine*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("Cv")
								$Organe = $Organe[$Organe.Length - 1]
								$Organe = $Organe.split(" ")
								$Organe = $Organe[0]
								$Organe = $Poste + "EIT" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION CV NON ORIGINE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition CV NON ORIGINE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition Cv non intermédiaire
							("*Cv*non interm*diaire*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("Cv")
								$Organe = $Organe[$Organe.Length - 1]
								$Organe = $Organe.split(" ")
								$Organe = $Organe[0]
								$Organe = $Poste + "EIT" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_INACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION CV NON INTERMEDIAIRE"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition CV NON INTERMEDIAIRE ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition FNITO
							("*FNITO*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("FNITO")
								$Organe = $Organe[$Organe.Length -1]
								$Organe = $Organe.replace("]","")
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "AXAUE" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION FNITO"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition FNITO ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
							
							# S'il y a une condition FNITD
							("*FNITD*") {
								$Organe = $AllTerms[$n]
								$Organe = $Organe.split("FNITD")
								$Organe = $Organe[$Organe.Length -1]
								$Organe = $Organe.replace("]","")
								$Organe = $Organe.replace(" ","")
								$Organe = $Organe.ToUpper()
								$Organe = $Poste + "AXAUS" + $Organe
								(	(""),
									("		SetIncident(""" + $Organe + """,""CONTROL_STATE_TO_ACTIVE"",""ACTION SUR L'ORGANE " + $Organe + " POUR LA CONDITION FNITD"");"),
									("		delay(2);")
								) >> $OutputScript
								'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition FNITD ecrite pour ' + $Instance + ' dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
							}
						}
					}
				}
#>
				# traitement sur l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Traitement du cas pour les itinéraires particuliers
				if($AuSettingUpTerms[$i]."Renvoi" -eq "(10)"){
					$Instance = $Instance + " 0"
					[Bool] $ItiAuPart = $TRUE
				}
				else {
					[Bool] $ItiAuPart = $FALSE
				}
				
				break
				
			}
		}
		
		# traitement sur l'instance
		$Instance = $Instance.replace(" par ","/")
		$Instance = $Instance.replace("-"," ")
		
		# Concatenation pour la clé complète
		$Key = $Key + " " + $Instance
		
		# Ecriture de l'action de clé dans le script
		(	(""),
			("		int keyTest = commandLineKey(""$Key"");"),
			("		int confirmSetUpTest = 0;")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Cle ' + $Key + ' ecrite dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
		
		# Ecriture de l'action de confirmation des itinéraires ou autorisations particuliers
		if ($ItiAuPart -eq $TRUE) {
			(	(""),
				("		int confirmSetUpTest = confirmSetUp();")
			) >> $OutputScript
			
			'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : Condition de confirmation de formation particuliere pour ' + $Instance + ' ecrite dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
		}
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeySettingUp : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'ecrire les action pour la destruction d'itinéraire ou pour reprendre une autorisation
# Utilisée pour les clé 24 et 38.1
Function WriteKeyDest {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Clé à rentrer dans le bandeau mode expert
		[Parameter(Position=0)]
		[String]
		$Key,
		# Instance pour la construction de la clé
		# Utilisée pour reconnaitre l'itinéraire ou l'autorisation en question
		[Parameter(Position=1)]
		[String]
		$Instance,
		# Chemin complet du Script de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
			
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de Key
		if ($Key -eq $null -or $Key -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : l''entree Key est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de Instance
		if ($Instance -eq $null -or $Instance -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : l''entree Instance est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		[bool] $break = $FALSE
		
		[Bool] $NotEPA = $TRUE
		
		# Test de condition de destruction type EPA
		# Boucle parcourant l'objet comportant les conditions
		for ([int] $i = 0; $i -lt $ItiDestEPA.count; $i++){
			# Test de correspondance des instances
			if ($ItiDestEPA[$i]."Itineraire" -eq $Instance) {
				# Transformation du delai en secondes
				[Int] $Delay = $ItiDestEPA[$i]."Temporisation (min)"
				$Delay = $Delay * 60 + 2
				# Traitement de l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Concatenation pour la clé complète
				$Key = $Key + " " + $Instance
				
				# Ecriture des actions de clé dans le script
				(	(""),
					("		// CLE 25 SUR LE CARRE " + $ItiDestEPA[$i]."Signal de depart"),
					("		int keyTest1 = commandLineKey(""25 " + $ItiDestEPA[$i]."Signal de depart" + """);"),
					("		// CLE $Key"),
					("		int keyTest2 = commandLineKey(""$Key"");"),
					("		// PRISE EN COMPTE DE LA TEMPORISATION"),
					("		delay(" + $Delay + ");"),
					("		// CLE $Key"),
					("		int keyTest3 = commandLineKey(""$Key"");"),
					("		int keyTest = keyTest1 + keyTest2 + keyTest3;")
				) >> $OutputScript
				
				$NotEPA = $FALSE
				
				'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : Cle ' + $Key + ' ecrite dans ' + $OutputScript + ' avec EPA'| Out-File -FilePath $LogFullPath -Append
				
				$break = $TRUE
				break
				
			}
		}

		[Bool] $NotDMT = $TRUE
		
		# Test de condition de destruction type DMT
		# Boucle parcourant l'objet comportant les conditions
		for ([int] $i = 0; $i -lt $ItiDestDMT.count -and $break -eq $TRUE; $i++){
			# Test de correspondance des instances
			if ($ItiDestDMT[$i]."Itineraire" -eq $Instance) {
				# Transformation du delai en secondes
				[Int] $Delay = $ItiDestDMT[$i]."Temporisation (min)"
				$Delay = $Delay * 60 + 2
				# Traitement de l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Concatenation pour la clé complète
				$Key = $Key + " " + $Instance
				
				# Ecriture de l'action de clé dans le script
				(	(""),
					("		// CLE $Key"),
					("		int keyTest = commandLineKey(""$Key"");"),
					("		// PRISE EN COMPTE DE LA TEMPORISATION"),
					("		delay(" + $Delay + ");")
				) >> $OutputScript
				
				$NotDMT = $FALSE
				
				'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : Cle ' + $Key + ' ecrite dans ' + $OutputScript + ' avec DMT'| Out-File -FilePath $LogFullPath -Append
				
				$break = $TRUE
				break
				
			}
				
		}
		
		[Bool] $NotSpe = $TRUE
		
		# Test de condition de destructions particulières
		# Boucle parcourant l'objet comportant les conditions
		for ([int] $i = 0; $i -lt $ItiDestSpeTerms.count -and $break -eq $TRUE; $i++){
			# Test de correspondance des instances
			if ($ItiDestSpeTerms[$i]."Itineraire" -eq $Instance) {
				# Traitement de l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Concatenation pour la clé complète
				$Key = $Key + " " + $Instance
				
				# Ecriture de l'action de clé dans le script
				(	("	// WARNING : CONDITION DESTRUCTION SPECIFIQUE"),
					(""),
					("		println(""ITINERAIRE A ENLEVER DES TESTS AUTOMATIQUES CAR SOUMIS A DES CONDITIONS PARTICULIERES"");"),
					("		int keyTest = commandLineKey(""$Key"");"),
					("")
				) >> $OutputScript
				
				$NotSpe = $FALSE
				
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : CONDITION PARTICULIERE DE DESTRUCTION POUR ' + $Key + ' A COMPLETER DANS ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
				
				$break = $TRUE
				break
				
			}
		}
		
		[Bool] $NotAu = $TRUE
		
		# Test de condition de destruction pour les autorisations
		# Boucle parcourant l'objet comportant les conditions
		for ([int] $i = 0; $i -lt $AuDest.count -and $break -eq $TRUE; $i++) {
			# Test de correspondance des instances
			if ($AuDest[$i]."Autorisation" -eq $Instance) {
				# Traitement de l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Concatenation pour la clé complète
				$Key = $Key + " " + $Instance
				
# AUCUNES CONDITIONS SPECIFIQUES SUR LE DOMAINE ATS
				
				# Ecriture de l'action de clé dans le script
				(	(""),
					("		// CLE $Key"),
					("		int keyTest = commandLineKey(""$Key"");")
				) >> $OutputScript
				
				$NotAu = $FALSE
				
				'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : Cle ' + $Key + ' ecrite dans ' + $OutputScript + ' avec DMT'| Out-File -FilePath $LogFullPath -Append
				
				$break = $TRUE
				break
				
			}
		}
		
		# Si c'est un itinéraire sans condition spécifique
		if (($NotEPA -eq $TRUE) -and ($NotDMT -eq $TRUE) -and ($NotSpe -eq $TRUE) -and ($NotAu -eq $TRUE)) {
				# Traitement de l'instance
				$Instance = $Instance.replace(" par ","/")
				$Instance = $Instance.replace("-"," ")
				
				# Concatenation pour la clé complète
				$Key = $Key + " " + $Instance
				
				# Ecriture de l'action de clé dans le script
				(	(""),
					("		// CLE $Key"),
					("		int keyTest = commandLineKey(""$Key"");")
				) >> $OutputScript
				
				'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : Cle ' + $Instance + ' ecrite dans ' + $OutputScript| Out-File -FilePath $LogFullPath -Append
		}
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKeyDest : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    }
}

# Fonction permetant d'ecrire l'action d'une clé dans un script + ajouter le test  des parametres
Function WriteKey {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Cle a rentrer dans le bandeau mode expert
		[Parameter(Position=0)]
		[String]
		$Key,
		# Indicateur pour le retour entier du test de cle
		[Parameter(Position=1)]
		[String]
		$i,
		# Chemin complet du Script de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de Key
		if ($Key -eq "" -or $Key -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : l''entree Key est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de i
		if ($i -eq "" -or $i -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : l''entree i est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour appeler la cle dans le script
		(	(""),
			("		// CLE $Key"),
			("		int keyTest$i = commandLineKey(""$Key"");")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteKey : Cle ' + $Key + ' ecrite dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Voir variable globale dans resultat
# Fonction permetant d'ecrire le debut du manager d'une famille de tests
Function WriteManagerStart {
	# Initialisation des parametres d'entree
	PARAM (
		# Identifiant du test generique
		[Parameter(Position=0)]
		[String]
		$testGenere,
		# Poste concerne
		[Parameter(Position=1)]
		[String]
		$Poste,
		# Chemin complet du manager de sortie
		[Parameter(Position=2)]
		[String]
		$OutputManager,
		# Documents d'entrees
		[Parameter(Position=3)]
		$EntryDocuments,
		# Chemin racine du rapport
		[Parameter(Position=4)]
		[String]
		$OutputPathReport,
		# Chemin complet du log de generation
		[Parameter(Position=5)]
		[String]
		$LogFullPath,
		# Chemin parent des scripts pour le test et le poste en question
		[Parameter(Position=6)]
		[String]
		$SpsDirectory
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputManager
		if ($OutputManager -eq "" -or $OutputManager -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''entree OutputManager est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de testGenere
		if ($testGenere -eq "" -or $testGenere -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''entree testGenere est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de Poste
		if ($Poste -eq "" -or $Poste -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''entree Poste est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de OutputPathReport
		if ($OutputPathReport -eq "" -or $OutputPathReport -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''entree OutputPathReport est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de SpsDirectory
		if ($SpsDirectory -eq "" -or $SpsDirectory -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''entree SpsDirectory est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la presentation dans le script de manager
		(	("//================================================================================================================="),
			("// Copyright (c) Siemens S.A.S. " + $(Get-Date)),
			("// All Rights Reserved, Confidential"),
			("//"),
			("// Identifiant du test generique : NExTEO_ATS_SFE_TST_PSO-" + $testGenere),
			("// Poste       : " + $Poste),
			("//")
		) > $OutputManager

		# Ecriture de l'ensemble des documents d'entree presentes dans la FE
		# Test du contenu de l'objet EntryDocuments
		if ($EntryDocuments -eq $null) {
			# Gestion de l'erreur objet document d'entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''objet contenants les documents d''entrees est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test des proprietes de l'objet EntryDocuments
		$EntryDocumentsTest = $EntryDocuments | Get-Member -MemberType NoteProperty | % {"$($_.Name)"}
		
		# Test de l'existence de "Titre"
		if ($EntryDocumentsTest -NotContains "Titre") {
			# Gestion de l'erreur objet ne contenant pas la propriete Titre
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''objet ne contient pas la propriete "Titre"' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test de l'existence de "Reference"
		if ($EntryDocumentsTest -NotContains "Reference") {
			# Gestion de l'erreur objet ne contenant pas la propriete Reference
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''objet ne contient pas la propriete "Reference"' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test de l'existence de "Version"
		if ($EntryDocumentsTest -NotContains "Version") {
			# Gestion de l'erreur objet ne contenant pas la propriete Version
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : l''objet ne contient pas la propriete "Version"' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de toutes les proprietes pour chaque documents d'entrees et ecriture des documents d'entrees dans le manager
		For ($n=1;$EntryDocuments[$n-1] -ne $null;$n++) {
			# Test du contenu de la propriete "Titre" pour le document n
			if ($EntryDocuments[$n-1].Titre -eq $null){
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : le document d''entree ' + $n + ' n''a pas de titre' | Out-File -FilePath $LogFullPath -Append
			}
			
			# Test du contenu de la propriete "Reference" pour le document n
			if ($EntryDocuments[$n-1].Reference -eq $null){
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : le document d''entree ' + $n + ' n''a pas de reference' | Out-File -FilePath $LogFullPath -Append
			}
			
			# Test du contenu de la propriete "Version" pour le document n
			if ($EntryDocuments[$n-1].Version -eq $null){
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : le document d''entree ' + $n + ' n''a pas de version' | Out-File -FilePath $LogFullPath -Append
			}
			
			# Ecriture du document d'entree n avec toutes ses proprietes dans le manager 
			("// Document d'entree " + $n + " : " + $EntryDocuments[$n-1].Titre + " \ " + $EntryDocuments[$n-1].Reference + " \ " + $EntryDocuments[$n-1].Version ) >> $OutputManager
		}

		# Ecriture du debut du scenario avec l'emplacement du rapport
		(	("//================================================================================================================="),
			(""),
			("void MANAGER_TST_PSO_P" + $Poste + "_" + $testGenere + "() {"),
			("	// INITIALISATION"),
			("	use system;"),
			("	variant GlobalResult = new Test;"),
			(""),
			("	// APPEL DE LA BOITE DE DIALOGUE POUR LE CHOIX D'EXECUTION"),
			("	int replayTest = callReplay(""$testGenere"",""$Poste"");"),
			(""),
			("	if (replayTest == 0) {"),
			(""),
			("		// FORMATAGE DE LA DATE POUR LA CREATION DU DOSSIER"),
			("		date h = h.now();"),
			("		string hexe = h.asString();"),
			("		hexe = hexe.replace(""/"",""_"");"),
			("		hexe = hexe.replace("" "",""_"");"),
			("		hexe = hexe.replace("":"",""_"");"),
			("		system.defineGlobal(""hexe"",hexe);"),
			(""),
			("		// CREATION DES VARIABLES NECESSAIRE POUR DETERMINER LE CHOIX D'EXECUTION"),
			("		int start = 0;"),
			("		int end = 0;"),
			("		bool whole = true;"),
			(""),
			("		GlobalResult.scenario(""MANAGER_TST_PSO_P" + $Poste + "_" + $testGenere + ".sps"",""NExTEO_ATS_SFE_TST_PSO-" + $testGenere + """);"),
			("		GlobalResult.start("""+ $OutputPathReport + "\TEST_ENTIER_"" + hexe + ""\01_RESULT\00_MANAGER\R_TST_PSO_P" + $Poste + "_" + $testGenere + ".result"");"),
			(""),
			("		// DEFINITION DES REPERTOIRES EN VARIABLES GLOBALES"),
			("		string scriptResultPath = ""$OutputPathReport\TEST_ENTIER_"" + hexe + ""\01_RESULT\01_SCRIPT\R_"";"),
			("		system.defineGlobal(""scriptResultPath"",scriptResultPath);"),
			(""),
			("		string LogExeFullPath = ""$OutputPathReport\TEST_ENTIER_"" + hexe + ""\00_LOG\Log_Exe_Entier_"" + hexe + ""_.txt"";"),
			("		system.defineGlobal(""LogExeFullPath"",LogExeFullPath);"),
			(""),
			("		string ScreenShotPath = ""$OutputPathReport\TEST_ENTIER_"" + hexe + ""\02_IMPR_ECRAN"";"),
			("		system.defineGlobal(""ScreenShotPath"",ScreenShotPath);"),
			(""),
			("		println(""DEBUT DU SUPER SCENARIO POUR LE DEROULE DES TESTS "+ $testGenere + " APPLIQUE AU POSTE " + $Poste + """);"),
			(""),
			("	}"),
			("	if (replayTest == 1) {"),
			(""),
			("		// APPEL DE LA BOITE DE DIALOGUE POUR LE CHOIX DE DEBUT D'EXECUTION"),
			("		int start = callStartChoice(""$SpsDirectory"");"),
			(""),
			("		if (start == -1) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""IMPOSSIBLE DE TROUVER LES SCRIPTS DE TESTS DANS L'EMPLACEMENT : $SpsDirectory"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		if (start == -2) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""CHOIX DANS LA BOITE DE DIALOGUE INCORRECT"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		if (start == -3) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""ERREUR EXCEPTIONNELLE"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		// APPEL DE LA BOITE DE DIALOGUE POUR LE CHOIX DE FIN D'EXECUTION"),
			("		int end = callEndChoice(""$SpsDirectory"");"),
			(""),
			("		if (end == -1) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""IMPOSSIBLE DE TROUVER LES SCRIPTS DE TESTS DANS L'EMPLACEMENT : $SpsDirectory"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		if (end == -2) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""CHOIX DANS LA BOITE DE DIALOGUE INCORRECT"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		if (end == -3) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""ERREUR EXCEPTIONNELLE"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		if (start > end) {"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("			println(""LE CHOIX DU NUMERO DE DEBUT ("" + start + "") EST PLUS GRAND QUE LE NUMERO DE FIN ("" + end + "")"");"),
			("			println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("			GlobalResult.fatalTest(false);"),
			("		}"),
			(""),
			("		// CREATION DE LA VARIABLE NECESSAIRE POUR DETERMINER LE CHOIX D'EXECUTION"),
			("		bool whole = false;"),
			(""),
			("		// FORMATAGE DE LA DATE POUR LA CREATION DU DOSSIER"),
			("		date h = h.now();"),
			("		string hexe = h.asString();"),
			("		hexe = hexe.replace(""/"",""_"");"),
			("		hexe = hexe.replace("" "",""_"");"),
			("		hexe = hexe.replace("":"",""_"");"),
			("		system.defineGlobal(""hexe"",hexe);"),
			(""),
			("		GlobalResult.scenario(""MANAGER_TST_PSO_P" + $Poste + "_" + $testGenere + ".sps"",""REJEU NExTEO_ATS_SFE_TST_PSO-" + $testGenere + """);"),
			("		GlobalResult.start("""+ $OutputPathReport + "\REJEU_"" + hexe + ""\01_RESULT\00_MANAGER\R_TST_PSO_P" + $Poste + "_" + $testGenere + ".result"");"),
			(""),
			("		// DEFINITION DES REPERTOIRES EN VARIABLES GLOBALES"),
			("		string scriptResultPath = ""$OutputPathReport\REJEU_"" + hexe + ""\01_RESULT\01_SCRIPT\R_"";"),
			("		system.defineGlobal(""scriptResultPath"",scriptResultPath);"),
			(""),
			("		string LogExeFullPath = ""$OutputPathReport\REJEU_"" + hexe + ""\00_LOG\Log_Exe_Rejeu_"" + hexe + ""_.txt"";"),
			("		system.defineGlobal(""LogExeFullPath"",LogExeFullPath);"),
			(""),
			("		string ScreenShotPath = ""$OutputPathReport\REJEU_"" + hexe + ""\02_IMPR_ECRAN"";"),
			("		system.defineGlobal(""ScreenShotPath"",ScreenShotPath);"),
			(""),
			("		println(""DEBUT DU SCENARIO DE REJEU POUR LE TESTS "+ $testGenere + " APPLIQUE AU POSTE " + $Poste + """);"),
			("		println(""DEBUT : "" + start + "" FIN : "" + end);"),
			("	}"),
			(""),
			("	if (replayTest == 2) {"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("		println(""CHOIX DANS LA BOITE DE DIALOGUE INCORRECT"");"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("		GlobalResult.fatalTest(false);"),
			("	}"),
			(""),
			("	if (replayTest == 3) {"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("		println(""ERREUR EXCEPTIONNELLE"");"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("		GlobalResult.fatalTest(false);"),
			("	}")
		) >> $OutputManager
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : Debut du manager de la famille de test ' + $testGenere + ' ecrit pour le poste ' + $Poste + ' dans ' + $OutputManager | Out-File -FilePath $LogFullPath -Append
	
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerStart : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permetant d'ecrire l'appel du script i depuis le manager d'une famille de tests
Function WriteManagerScript {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Numero du script de test
		[Parameter(Position=0)]
		[String]
		$numTest,
		# Test specifique
		[Parameter(Position=1)]
		[String]
		$TestSpe,
		# Instance complete
		[Parameter(Position=2)]
		[String]
		$InstanceCell,
		# Chemin complet du manager de sortie
		[Parameter(Position=3)]
		[String]
		$OutputManager,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputManager
		if ($OutputManager -eq "" -or $OutputManager -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : l''entree OutputManager est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de TestSpe
		if ($TestSpe -eq "" -or $TestSpe -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : l''entree TestSpe est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de numTest
		if ($numTest -eq "" -or $numTest -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : l''entree numTest est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de InstanceCell
		if ($InstanceCell -eq "" -or $InstanceCell -eq $Null) {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : l''entree InstanceCell est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		
		# Ecriture de la ligne de code pour appeler le script i depuis le manager
		(	(""),
			("	// APPEL DU SCRIPT DE TEST " + $TestSpe),
			("	bool ReplayExec = (" + $numTest + " <= end && " + $numTest + " >= start);"),
			("	if (whole == true || ReplayExec == true) {"),
			("		bool Tg" + $numTest + " = " + $TestSpe +"();"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("		println(""DEBUT DU TEST : " + $TestSpe + """);"),
			("		println(""APPLIQUE A L'INSTANCE :" + $InstanceCell + """);"),
			("		GlobalResult.test(Tg" + $numTest + ",""RESULTAT TEST " + $TestSpe + " : OK"",""RESULTAT TEST " + $TestSpe + " : NOK"");"),
			("		println(""FIN DU TEST : " + $TestSpe + """);"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			("	}")
		) >> $OutputManager
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : Appel du script ' + $TestSpe + ' ecrit dans ' + $OutputManager | Out-File -FilePath $LogFullPath -Append
	
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerScript : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permetant d'ecrire la fin du manager d'une famille de tests
Function WriteManagerEnd {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Identifiant du test generique
		[Parameter(Position=0)]
		[String]
		$testGenere,
		# Poste concerne
		[Parameter(Position=1)]
		[String]
		$Poste,
		# Chemin complet du manager de sortie
		[Parameter(Position=2)]
		[String]
		$OutputManager,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerEnd : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputManager
		if ($OutputManager -eq "" -or $OutputManager -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerEnd : l''entree OutputManager est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de testGenere
		if ($testGenere -eq "" -or $testGenere -eq $Null) {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteManagerEnd : l''entree testGenere est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de Poste
		if ($Poste -eq "" -or $Poste -eq $Null) {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteManagerEnd : l''entree Poste est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Ecriture de la ligne de code pour ecrire la fin du manager
		(	(""),
			("	if (replayTest == 0) {"),
			("		println(""FIN DU SUPER SCENARIO POUR LE DEROULE DE TOUS LES TESTS " + $testGenere + " APPLIQUE AU POSTE " + $Poste + """);"),
			("	}"),
			(""),
			("	if (replayTest == 1) {"),
			("		println(""FIN DU SCENARIO DE REJEU POUR LES TESTS " + $testGenere + " APPLIQUE AU POSTE " + $Poste + """);"),
			("	}"),
			("}")
		) >> $OutputManager
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteManagerEnd : FIN DE GENERATION DU MANAGER POUR LE TEST ' + $testGenere + ' CONCERNANT LE POSTE ' + $Poste| Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteManagerEnd : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permetant d'ecrire le debut du script d'une famille de tests
Function WriteScriptStart {
	# Initialisation des parametres d'entree
	PARAM (
		# Identifiant du test generique
		[Parameter(Position=0)]
		[String]
		$IdTestGen,
		# Identifiant du test specifique
		[Parameter(Position=1)]
		[String]
		$IdTestSpe,
		# nom du test et de l'onglet
		[Parameter(Position=2)]
		[String]
		$TestSpe,
		# numero du test specifique
		[Parameter(Position=3)]
		[String]
		$numTest,
		# Instance complete
		[Parameter(Position=4)]
		[String]
		$InstanceCell,
		# Nom du test
		[Parameter(Position=5)]
		[String]
		$TestName,
		# Moyen d'essai
		[Parameter(Position=6)]
		[String]
		$TestMedium,
		# Version de l'ATS
		[Parameter(Position=7)]
		[String]
		$ATSVersion,
		# Poste concerne
		[Parameter(Position=8)]
		[String]
		$Poste,
		# Chemin complet du script de sortie
		[Parameter(Position=9)]
		[String]
		$OutputScript,
		# Documents d'entrees
		[Parameter(Position=10)]
		$EntryDocuments,
		# Chemin complet du log de generation
		[Parameter(Position=11)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
	
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq $null -or $LogFullPath -eq "") {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de IdTestGen
		if ($IdTestGen -eq $null -or $IdTestGen -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree IdTestGen est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de IdTestSpe
		if ($IdTestSpe -eq $null -or $IdTestSpe -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree IdTestSpe est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de TestSpe
		if ($TestSpe -eq $null -or $TestSpe -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree TestSpe est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de numTest
		if ($numTest -eq $null -or $numTest -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree numTest est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de InstanceCell
		if ($InstanceCell -eq $null -or $InstanceCell -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree InstanceCell est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de TestName
		if ($TestName -eq $null -or $TestName -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree TestName est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de TestMedium
		if ($TestMedium -eq $null -or $TestMedium -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree TestMedium est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de ATSVersion
		if ($ATSVersion -eq $null -or $ATSVersion -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree ATSVersion est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de Poste
		if ($Poste -eq $null -or $Poste -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree Poste est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la presentation dans le script
		(	("//================================================================================================================="),
			("// Copyright (c) Siemens S.A.S. " + $(Get-Date)),
			("// All Rights Reserved, Confidential"),
			("//"),
			("// Identifiant du test generique : " + $IdTestGen),
			("// Identifiant du test specifique : " + $IdTestSpe),
			("// Instance : " + $InstanceCell),
			("// Nom : " + $TestName),
			("//"),
			("// Moyen d'essai : " + $TestMedium),
			("// Version de l'ATS : " + $ATSVersion),
			("//"),
			("// Poste       : " + $Poste),
			("//")
		) > $OutputScript

		# Ecriture de l'ensemble des documents d'entree presentes dans la FE
		# Test du contenu de l'objet EntryDocuments
		if ($EntryDocuments -eq $null) {
			# Gestion de l'erreur objet document d'entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''objet contenants les documents d''entrees est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test des proprietes de l'objet EntryDocuments
		$EntryDocumentsTest = $EntryDocuments | Get-Member -MemberType NoteProperty | % {"$($_.Name)"}
		
		# Test de l'existence de "Titre"
		if ($EntryDocumentsTest -NotContains "Titre") {
			# Gestion de l'erreur objet ne contenant pas la propriete Titre
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''objet ne contient pas la propriete "Titre"' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test de l'existence de "Reference"
		if ($EntryDocumentsTest -NotContains "Reference") {
			# Gestion de l'erreur objet ne contenant pas la propriete Reference
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''objet ne contient pas la propriete "Reference"' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test de l'existence de "Version"
		if ($EntryDocumentsTest -NotContains "Version") {
			# Gestion de l'erreur objet ne contenant pas la propriete Version
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : l''objet ne contient pas la propriete "Version"' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de toutes les proprietes pour chaque documents d'entrees et ecriture des documents d'entrees dans le manager
		For ($n=1;$EntryDocuments[$n-1] -ne $null;$n++) {
			# Test du contenu de la propriete "Titre" pour le document n
			if ($EntryDocuments[$n-1].Titre -eq $null){
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : le document d''entree ' + $n + ' n''a pas de titre' | Out-File -FilePath $LogFullPath -Append
			}
			
			# Test du contenu de la propriete "Reference" pour le document n
			if ($EntryDocuments[$n-1].Reference -eq $null){
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : le document d''entree ' + $n + ' n''a pas de reference' | Out-File -FilePath $LogFullPath -Append
			}
			
			# Test du contenu de la propriete "Version" pour le document n
			if ($EntryDocuments[$n-1].Version -eq $null){
				'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : le document d''entree ' + $n + ' n''a pas de version' | Out-File -FilePath $LogFullPath -Append
			}
			
			# Ecriture du document d'entree n avec toutes ses proprietes dans le manager 
			("// Document d'entree " + $n + " : " + $EntryDocuments[$n-1].Titre + " \ " + $EntryDocuments[$n-1].Reference + " \ " + $EntryDocuments[$n-1].Version ) >> $OutputScript
		}

		# Ecriture du debut du script
		(	("//================================================================================================================="),
			(""),
			("bool " + $TestSpe +"() {"),
			("	use system;"),
			("	use S_CCK_MCMD_AVERTISSEMENT;"),
			("	use S_CCK_ALARME;"),
			("	use ALARM_COLLECTOR;"),
			(""),
			("	variant T" + $numTest + " = new Test;"),
			(""),
			("	bool Resultat = true;"),
			(""),
			("	T" + $numTest +".scenario(""" + $TestSpe + ".sps""" + ",""" + "Test " + $IdTestGen + " applique a l'instance " + $InstanceCell + """);"),
			("	string scriptResultPath = system.getGlobal(""scriptResultPath"");"),
			("	T" + $numTest + ".start(scriptResultPath + ""$TestSpe.result"");"),
			("	println(endl + ""$IdTestSpe"");"),
			("	println(""Name : " + $IdTestSpe.Substring(15,$IdTestSpe.Length - 15) + """ + endl);")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : Debut du script ' + $IdTestSpe + ' ecrit pour le poste ' + $Poste + ' dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptStart : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'ecrire le test d'existence des Ts et des organes
Function WriteExistTest {
	# Initialisation des parametres d'entree
	PARAM (
		# Liste de toutes les Ts
		[Parameter(Position=0)]
		$TsList,
		# Liste de tous les organes
		[Parameter(Position=1)]
		$OrganList,
		# Liste de tous les appareils de voie
		[Parameter(Position=2)]
		$trackDeviceList,
		# Chemin complet du script de sortie
		[Parameter(Position=3)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteExistTest : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteExistTest : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de TsList
		if ($TsList -eq "" -or $TsList -eq $Null) {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteExistTest : l''entree TsList est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de OrganList
		if ($OrganList -eq "" -or $OrganList -eq $Null) {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteExistTest : l''entree OrganList est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de trackDeviceList
		if ($trackDeviceList -eq "" -or $trackDeviceList -eq $Null) {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteExistTest : l''entree trackDeviceList est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Ecriture de toutes les Ts et test de leur existence
		if ($TsList -ne "" -and $TsList -ne $NULL) {
			(	(""),
				("	//TEST D EXISTENCE DES TS"),
				("	println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
				("	println(""TEST D EXISTENCE DES TS"");"),
				("	bool tsG = true;")
			) >> $OutputScript
			
			For ([Int] $i = 0; $i -lt $TsList.Count; $i++) {
				[String] $tsName = $TsList[$i]
				(	("	bool ts$i = exist(""$tsName"");"),
					("	bool tsG = tsG && ts$i;"),
					("	if(ts$i == true){"),
					("		variant $tsName = get(""$tsName"");"),
					("	}")
				) >> $OutputScript
			}
			
			(	(""),
				("	if(tsG == false){"),
				("		println(""AU MOINS UNE TS N'EXISTE PAS"");"),
				("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
				("		return tsG;"),
				("	}"),
				(""),
				("	else{"),
				("		println(""TOUTES LES TS EXISTENT"");"),
				("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
				("	}")
			) >> $OutputScript
		
			'[INFO] - ' + $(Get-Date) + ' : Fonction WriteExistTest : test d''existence et utilisation des Ts ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
		}
		
		# Ecriture de tous les organes et test de leur existence
		if ($OrganList -ne "" -and $OrganList -ne $NULL) {
			(	(""),
				("	// TEST D EXISTENCE DES ORGANES"),
				("	println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
				("	println(""TEST D EXISTENCE DES ORGANES"");"),
				("	bool orgG = true;")
			) >> $OutputScript
			
			For ([Int] $i = 0; $i -lt $OrganList.Count; $i++) {
				[String] $organName = $OrganList[$i]
				(	("	bool org$i = OrganExist(""$organName"");"),
					("	bool orgG = orgG && org$i;")
				) >> $OutputScript
			}
			
			(	(""),
				("	if(orgG == false){"),
				("		println(""AU MOINS UN ORGANE N'EXISTE PAS"");"),
				("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
				("		return orgG;"),
				("	}"),
				(""),
				("	else{"),
				("		println(""TOUS LES ORGANES EXISTENT"");"),
				("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
				("	}")
			) >> $OutputScript
			
			'[INFO] - ' + $(Get-Date) + ' : Fonction WriteExistTest : test d''existence des organes ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
		}
		
		# Ecriture de tous les appareils de voie et test de leur existence
		if ($trackDeviceList -ne "" -and $trackDeviceList -ne $NULL) {
			(	(""),
				("	// TEST D EXISTENCE DES APPAREILS DE VOIE"),
				("	println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
				("	println(""TEST D EXISTENCE DES APPAREILS DE VOIE"");"),
				("	bool devG = true;")
			) >> $OutputScript
			
			For ([Int] $i = 0; $i -lt $trackDeviceList.Count; $i++) {
				[String] $deviceName = $trackDeviceList[$i]
				(	("	bool dev$i = exist(""$deviceName"");"),
					("	bool devG = devG && dev$i;")
				) >> $OutputScript
			}
			
			(	(""),
				("	if(devG == false){"),
				("		println(""AU MOINS UN APPAREIL DE VOIE N'EXISTE PAS"");"),
				("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
				("		return devG;"),
				("	}"),
				(""),
				("	else{"),
				("		println(""TOUS LES APPAREILS DE VOIE EXISTENT"");"),
				("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
				("	}")
			) >> $OutputScript
			
			'[INFO] - ' + $(Get-Date) + ' : Fonction WriteExistTest : test d''existence des appareils de voie ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
		}
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteExistTest : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'ecrire l'initialisation d'un script
Function WriteScriptInit {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Identifiant du test specifique
		[Parameter(Position=0)]
		[String]
		$IdTestSpe,
		# Poste
		[Parameter(Position=1)]
		[String]
		$Poste,
		# Chemin complet du manager de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
	
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq $null -or $LogFullPath -eq "") {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptInit : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptInit : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de IdTestSpe
		if ($IdTestSpe -eq $null -or $IdTestSpe -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptInit : l''entree IdTestSpe est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de Poste
		if ($Poste -eq $null -or $Poste -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptInit : l''entree Poste est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de l'initialisation du script
		(	(""),
			("	// INITIALISATION"),
			("	println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("	println(""INITIALISATION PROTECTIONS ET ITINERAIRES / AUTORISATIONS"");"),
			("	println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			(""),
			("	println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"");"),
			("	bool init1 = CTRLAllProt();"),
			("	bool init2 = CTRLAllFnitFnau(" + $Poste + ");"),
			("	bool initTot = init1 && init2;"),
			("	if (initTot){"),
			(""),
			("		println(""INITIALISATION OK DEBUT DU TEST"");"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);")
		) >> $OutputScript
	
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteScriptInit : Initialisation du script ' + $IdTestSpe + ' ecrit pour le poste ' + $Poste + ' dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptInit : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'ecrire la presentation d'un step dans un script d'une famille de tests
Function WriteStepPres {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Numero du step
		[Parameter(Position=0)]
		[String]
		$numStep,
		# Action du step
		[Parameter(Position=1)]
		[String]
		$displayStep,
		# Chemin complet du manager de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq $null -or $LogFullPath -eq "") {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepPres : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepPres : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de numStep
		if ($numStep -eq $null -or $numStep -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteStepPres : l''entree numStep est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de displayStep
		if ($displayStep -eq $null -or $displayStep -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteStepPres : l''entree displayStep est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Ecriture du debut d'un step dans le script
		(	(""),
			("	// STEP $numStep : $displayStep"),
			("		stepPres($numStep,""$displayStep"");")
		) >> $OutputScript
	
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteStepPres : presentation du step ' + $numStep + ' ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepPres : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'ecrire la fin d'un step dans un script d'une famille de tests
Function WriteStepEnd {
	# Initialisation des paramètres d'entrée
	PARAM (
		# expression booleenne pour le test
		[Parameter(Position=0)]
		[String]
		$exp,
		# Numero du step
		[Parameter(Position=1)]
		[String]
		$numStep,
		# Commentaire de test
		[Parameter(Position=2)]
		[String]
		$comment,
		# Numero pour le variant
		[Parameter(Position=3)]
		[String]
		$numTest,
		# Chemin complet du manager de sortie
		[Parameter(Position=4)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=5)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
	
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq $null -or $LogFullPath -eq "") {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de numStep
		if ($numStep -eq $null -or $numStep -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : l''entree numStep est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de comment
		if ($comment -eq $null -or $comment -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : l''entree comment est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de numTest
		if ($numTest -eq $null -or $numTest -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : l''entree numTest est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de exp
		if ($exp -eq $null -or $exp -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : l''entree exp est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la fin d'un step dans le script
		(	(""),
			("		bool Resultat$numStep = $exp;"),
			("		stepEnd(Resultat$numStep,T$numTest,$numStep,""$comment"");"),
			("		Resultat = Resultat && Resultat$numStep;")
		) >> $OutputScript
	
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteStepEnd : fin du step ' + $numStep + ' ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'ecrire la fin d'un script
Function WriteScriptEnd {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Test specifique
		[Parameter(Position=0)]
		[String]
		$TestSpe,
		# Poste
		[Parameter(Position=1)]
		[String]
		$Poste,
		# Chemin complet du manager de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
	
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq $null -or $LogFullPath -eq "") {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq $null -or $OutputScript -eq "") {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de TestSpe
		if ($TestSpe -eq $null -or $TestSpe -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : l''entree TestSpe est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Test du contenu de Poste
		if ($Poste -eq $null -or $Poste -eq "") {
			# Gestion de l'erreur entree vide
			'[WARNING] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : l''entree Poste est vide' | Out-File -FilePath $LogFullPath -Append
		}
		
		# Ecriture de l'initialisation du script
		(	("		return Resultat;"),
			("	}"),
			("	else{"),
			(""),
			("		println(""INITIALISATION NOK"");"),
			("		println(""----------------------------------------------------------------------------------------------------------------------------------------------------------------"" + endl);"),
			(""),
			("		return initTot;"),
			("	}"),
			("}")
		) >> $OutputScript
	
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : fin du script ' + $TestSpe + ' ecrit pour le poste ' + $Poste + ' dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScriptEnd : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permettant d'écrire une action à pied d'oeuvre, une commande ou un dérangement dans un script
Function WriteOrganeManualAction {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Nom de l'organe
		[Parameter(Position=0)]
		[String]
		$organName,
		# Commande à appliquer
		[Parameter(Position=1)]
		[String]
		$command,
		# Affichage dans le script
		[Parameter(Position=1)]
		[String]
		$display,
		# Chemin complet du Script de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de organName
		if ($organName -eq "" -or $organName -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : l''entree organName est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de command
		if ($command -eq "" -or $command -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : l''entree command est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		# Test du contenu de display
		if ($display -eq "" -or $display -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : l''entree display est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test de correspondance de la commande avec les commandes existantes
		$existingCommands = @("ACTION_CANCEL_TRANSIT","ACTION_GIVE_AUTHORIZATION","ACTION_TAKE_BACK_AUTHORIZATION","ACTION_GIVE_ENTRY_AUTHORIZATION","ACTION_TAKE_BACK_ENTRY_AUTHORIZATION","ACTION_GIVE_EXIT_AUTHORIZATION","ACTION_TAKE_BACK_EXIT_AUTHORIZATION","ACTION_READY_FOR_DEPARTURE","CONTROL_TO_LEFT","CONTROL_TO_RIGHT","CONTROL_STATE_TO_INACTIVE","CONTROL_STATE_TO_ACTIVE","CONTROL_STATE_TO_INCONSISTENT","CONTROL_STATE_TO_NOT_BOOSTED","BLOCK","UNBLOCK","INCIDENT_MAINTAIN_LEFT","INCIDENT_MAINTAIN_RIGHT","INCIDENT_UNCONTROL_LEFT","INCIDENT_UNCONTROL_RIGHT","INCIDENT_DESTRUCTION_FAILURE","INCIDENT_FORCE_OCCUPIED","INCIDENT_MAINTAIN_OCCUPIED","INCIDENT_FORCE_FREE","INCIDENT_FORCE_ACTIVE","INCIDENT_MAINTAIN_ACTIVE","3INCIDENT_FORCE_INACTIVE","INCIDENT_OPENING_FAILURE","INCIDENT_FORCE_FALL","RAISE_INCIDENT","PUT_OUT_OF_SERVICE","PUT_IN_SERVICE")

		[Bool] $commandExist = $FALSE 
		
		for ([int] $i = 0; $i -lt $existingCommands.count -and $commandExist -eq $FALSE; $i++){
			if ($existingCommands[$i] -eq $command) {
				$commandExist = $TRUE
			}
		}

		# Test de l'existance de la commande
		if ($commandExist -eq $FALSE) {
			# Gestion de l'erreur commande inconnue
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : la commande sur ' + $organName + ' est inconnue' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour effectuer l'action a pied d'oeuvre / commande dans le script
		(	(""),
			("		SetIncident(""$organName"",""$command"",""$display"");"),
			("		delay(0.5);")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : action ' + $command + ' pour l''organe ' + $organName +' ecrite dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteOrganeManualAction : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'un etat dans les fichiers de Log CCK
Function WriteLogVerif {
# Initialisation des paramètres d'entrée
	PARAM (
		# Nom de l'organe
		[Parameter(Position=0)]
		[String]
		$organName,
		# Etat de l'organe
		[Parameter(Position=1)]
		[String]
		$organState,
		# Horodatage a tester
		[Parameter(Position=2)]
		[String]
		$h,
		# Chemin complet du Script de sortie
		[Parameter(Position=3)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de organName
		if ($organName -eq "" -or $organName -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree organName est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de organState
		if ($organState -eq "" -or $organState -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree organState est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de h
		if ($h -eq "" -or $h -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree h est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture du test dans les Logs CCK dans le script
		(	(""),
			("		bool Verif$organName = testTraceLogOrganState(""$organName"",$organState,$h);")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : test d''etat de ' + $organName + ' dans les Logs CCK ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'un etat dans les fichiers de Log CCK avant une action
Function WriteLogVerifBeforeAction {
# Initialisation des paramètres d'entrée
	PARAM (
		# Nom de l'organe
		[Parameter(Position=0)]
		[String]
		$organName,
		# Etat de l'organe
		[Parameter(Position=1)]
		[String]
		$organState,
		# Horodatage a tester
		[Parameter(Position=2)]
		[String]
		$h,
		# Chemin complet du Script de sortie
		[Parameter(Position=3)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de organName
		if ($organName -eq "" -or $organName -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree organName est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de organState
		if ($organState -eq "" -or $organState -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree organState est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de h
		if ($h -eq "" -or $h -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : l''entree h est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture du test dans les Logs CCK dans le script
		(	(""),
			("		bool Verif$organName = testTraceLogOrganStateBeforeAction(""$organName"",$organState,$h);")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : test d''etat de ' + $organName + ' dans les Logs CCK ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteLogVerif : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'une alarme avec argument
Function WriteAlarmVerif1 {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Horodatage a tester
		[Parameter(Position=0)]
		[String]
		$h,
		# Identifiant de l'alarme
		[Parameter(Position=1)]
		[String]
		$idAlarm,
		# Argument de l'alarme
		[Parameter(Position=2)]
		[String]
		$Arg,
		# Chemin complet du Script de sortie
		[Parameter(Position=3)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de h
		if ($h -eq "" -or $h -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : l''entree h est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de idAlarm
		if ($idAlarm -eq "" -or $idAlarm -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : l''entree idAlarm est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de Arg
		if ($Arg -eq "" -or $Arg -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : l''entree Arg est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour effectuer le test sur l'alarme dans le script
		(	(""),
			("		bool testAlarmVerif1 = testAlarmIdArgH($h,""$idAlarm"",""$Arg"");")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : test de l''alarme ' + $idAlarm + ' avec l''argument ' + $Arg +' ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif1 : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'une alarme sans argument
Function WriteAlarmVerif2 {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Horodatage a tester
		[Parameter(Position=0)]
		[String]
		$h,
		# Identifiant de l'alarme
		[Parameter(Position=1)]
		[String]
		$idAlarm,
		# Chemin complet du Script de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif2 : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif2 : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de h
		if ($h -eq "" -or $h -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif2 : l''entree h est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de idAlarm
		if ($idAlarm -eq "" -or $idAlarm -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif2 : l''entree idAlarm est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour effectuer le test sur l'alarme dans le script
		(	(""),
			("		bool testAlarmVerif2 = testAlarmIdH($h,""$idAlarm"");")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif2 : test de l''alarme ' + $idAlarm + ' sans argument ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAlarmVerif2 : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'un avertissement avec commande
Function WriteAvertVerif1 {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Horodatage a tester
		[Parameter(Position=0)]
		[String]
		$h,
		# Identifiant de l'avertissement
		[Parameter(Position=1)]
		[String]
		$idAvert,
		# Identifiant de la commande
		[Parameter(Position=2)]
		[String]
		$idCommand,
		# Chemin complet du Script de sortie
		[Parameter(Position=3)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de h
		if ($h -eq "" -or $h -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : l''entree h est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de idAvert
		if ($idAvert -eq "" -or $idAvert -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : l''entree idAvert est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de idCommand
		if ($idCommand -eq "" -or $idCommand -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : l''entree idCommand est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour effectuer le test sur l'avertissement dans le script
		(	(""),
			("		bool testAvertVerif1 = testAvertIdCommandH($h,""$idAvert"",""$idCommand"");")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : test de l''avertissement ' + $idAvert + ' avec la commande ' + $idCommand +' ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif1 : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'un avertissement sans commande
Function WriteAvertVerif2 {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Horodatage a tester
		[Parameter(Position=0)]
		[String]
		$h,
		# Identifiant de l'avertissement
		[Parameter(Position=1)]
		[String]
		$idAvert,
		# Chemin complet du Script de sortie
		[Parameter(Position=2)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif2 : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif2 : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de h
		if ($h -eq "" -or $h -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif2 : l''entree h est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de idAvert
		if ($idAvert -eq "" -or $idAvert -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif2 : l''entree idAvert est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour effectuer le test sur l'avertissement dans le script
		(	(""),
			("		bool testAvertVerif2 = testAvertIdH($h,""$idAvert"");")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif2 : test de l''avertissement ' + $idAvert + ' sans commande ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteAvertVerif2 : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Test booléen de vérification d'une TS
Function WriteTSVerif {
	# Initialisation des paramètres d'entrée
	PARAM (
		# Variant pour la Ts testé
		[Parameter(Position=0)]
		[String]
		$ts,
		# Etat de la Ts testé
		[Parameter(Position=1)]
		[String]
		$tsState,
		# Valeur de la Ts testé
		[Parameter(Position=2)]
		[String]
		$tsValue,
		# Chemin complet du Script de sortie
		[Parameter(Position=3)]
		[String]
		$OutputScript,
		# Chemin complet du log de generation
		[Parameter(Position=4)]
		[String]
		$LogFullPath
	)

	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de ts
		if ($ts -eq "" -or $ts -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : l''entree ts est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de tsState
		if ($tsState -eq "" -or $tsState -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : l''entree tsState est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de tsValue
		if ($tsValue -eq "" -or $tsValue -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : l''entree tsValue est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de code pour effectuer le test sur la ts dans le script
		(	(""),
			("		bool Verif$ts = testTS($ts,""$tsState"",$tsValue);")
		) >> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : test de la TS ' + $ts + ' a l''etat ' + $tsState + ' et a la valeur ' + $tsValue + ' ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	}
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteTSVerif : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
    
	}
}

# Fonction permetant d'ecrire l'impression ecran a effectuer dans le script de test
Function WriteScreenShot {
	# Initialisation des parametres d'entree
	PARAM (
		# Nom de l'impression ecran
		[Parameter(Position=0)]
		[String]
		$ScreenShotName,
		# Chemin complet du Script de sortie
		[Parameter(Position=1)]
		[String]
		$OutputScript,
		# Nom de la fenetre a mettre en avant
		[Parameter(Position=2)]
		$WindowName,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScreenShot : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de ScreenShotName
		if ($ScreenShotName -eq "" -or $ScreenShotName -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScreenShot : l''entree ScreenShotName est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de OutputScript
		if ($OutputScript -eq "" -or $OutputScript -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : l''entree OutputScript est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de WindowName
		if ($WindowName -eq "" -or $WindowName -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteKey : l''entree WindowName est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
	
		# Ecriture de la ligne de commande dans le script
		(	(""),
			("		commandLineScreenshot(""$ScreenShotName"",""$WindowName"");")
		)>> $OutputScript
		
		'[INFO] - ' + $(Get-Date) + ' : ScreenShot ' + $ScreenShotName + ' pour la fenetre ' + $WindowName + ' ecrit dans ' + $OutputScript | Out-File -FilePath $LogFullPath -Append
	
	}
	
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteScreenShot : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
		
    }
}

# Fonction permetant de creer le fichier sps d'initialisation des variables d'environnement
Function WriteInitAllEnvVar {
	# Initialisation des parametres d'entree
	PARAM (
		# Chemin parent des scripts pour le test et le poste en question
		[Parameter(Position=0)]
		[String]
		$SpsDirectory,
		# Chemin complet du Script de sortie
		[Parameter(Position=1)]
		[String]
		$OutputInitAllEnvVar,
		# Chemin complet du log de generation
		[Parameter(Position=3)]
		[String]
		$LogFullPath
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		# Test du contenu de LogFullPath
		if ($LogFullPath -eq "" -or $LogFullPath -eq $Null) {
			# Gestion de l'erreur entree vide
			[String] $ErrorMsg = '[ERROR] - ' + $(Get-Date) + ' : Fonction WriteInitAllEnvVar : l''entree LogFullPath est vide'
			write-host $ErrorMsg
			start-sleep 2
			exit
		}
		
		# Test du contenu de SpsDirectory
		if ($SpsDirectory -eq "" -or $SpsDirectory -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteInitAllEnvVar : l''entree SpsDirectory est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Test du contenu de OutputInitAllEnvVar
		if ($OutputInitAllEnvVar -eq "" -or $OutputInitAllEnvVar -eq $Null) {
			# Gestion de l'erreur entree vide
			'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteInitAllEnvVar : l''entree OutputInitAllEnvVar est vide' | Out-File -FilePath $LogFullPath -Append
			exit
		}
		
		# Ecriture de la ligne de commande dans le script		
		(	("//================================================================================================================="),
			("// Copyright (c) Siemens S.A.S. " + $(Get-Date)),
			("// All Rights Reserved, Confidential"),
			("// Author	: Guillaume POIRIER"),
			("//"),
			("// Version	: 0.1"),
			("// File		: INIT_ALL_ENV_VAR"),
			("// Content	: Initialise les variables d'environement pour les programmes de type ""utilitaires"""),
			("//================================================================================================================="),
			("void INIT_ALL_ENV_VAR() {"),
			(""),
			("	use system;"),
			("	string INPUT_SPS = ""D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\00_INPUT_SPS"";"),
			("	string SCRIPT = ""$SpsDirectory"";"),
			("	string SPOCC_SEQ = system.getEnvironmentVariable(""SPOCC_SEQ"");"),
			("	system.setEnvironmentVariable(""SPOCC_SEQ"", SPOCC_SEQ + "";"" + INPUT_SPS + "";"" + SCRIPT + "";"");"),
			("	println(""Initialisation des variables d'environnement ok"");"),
			("}")
		) > $OutputInitAllEnvVar
		
		'[INFO] - ' + $(Get-Date) + ' : Fonction WriteInitAllEnvVar : initialisation des variables d''environement ecrit dans ' + $OutputInitAllEnvVar + ' avec ' + $SpsDirectory + ' comme repertoire pour les scripts appeles' | Out-File -FilePath $LogFullPath -Append
	
	}
	
	catch {
		
		# Gestion des exeptions
		'[ERROR] - ' + $(Get-Date) + ' : Fonction WriteInitAllEnvVar : ' + $($_.exception.message) | Out-File -FilePath $LogFullPath -Append
		
    }
}

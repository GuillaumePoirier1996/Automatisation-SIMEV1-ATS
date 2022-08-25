#####################################################################################################################
# Copyright (c) Siemens S.A.S. 06/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	1.0
# Description 	: 	Manager de generateur permettant de choisir le poste concerné, le générateur de famille de test,
#					la BDD et la FE et de creer les repertoire associés
#
# README 		:	lire la documentation associée
#####################################################################################################################
# INITIALISATION
PARAM (
	# Chemin pour les generateurs de tests
	[String] $generatorDirectory = "D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\02_GENERATEUR",
	# Chaine de caractere a reconnaitre dans le titre
	[String] $filenamePattern = "Generateur_TST_PSO_",
	# Chemin racine pour la version du code
	[String] $RootPath = "D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1",
	# Chemin Complet pour le dossier XX_PXX
	[String] $XX_PXX = "$RootPath\XX_PXX"
)

try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'
		
		write-host "--------------------------------------------------------------------------------"
		write-host "DEBUT DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
		write-host "--------------------------------------------------------------------------------"

		# CHOIX DU GENERATEUR

		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing

		#####-- Boite de dialogue
		$form = New-Object System.Windows.Forms.Form
		$form.Text = 'Choix du Poste'
		$form.Size = New-Object System.Drawing.Size(300,300)
		$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
		$form.StartPosition = 'CenterScreen'
		$form.BackColor = "Black"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$form.MaximizeBox = $False

		#####-- Bouton OK
		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Point(10,220)
		$OKButton.Size = New-Object System.Drawing.Size(100,30)
		$OKButton.Text = 'OK'
		$OKButton.BackColor = "Green"
		$OKButton.ForeColor = "White"
		$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $OKButton
		$form.Controls.Add($OKButton)

		#####-- Bouton Cancel
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Point(175,220)
		$CancelButton.Size = New-Object System.Drawing.Size(100,30)
		$CancelButton.Text = 'Cancel'
		$CancelButton.BackColor = "Red"
		$CancelButton.ForeColor = "White"
		$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $CancelButton
		$form.Controls.Add($CancelButton)

		#####-- Message
		$label = New-Object System.Windows.Forms.Label
		$label.Location = New-Object System.Drawing.Point(10,10)
		$label.Size = New-Object System.Drawing.Size(265,30)
		$label.Text = "Selectionnez le generateur de votre famille de test : "
		$label.ForeColor = "White"
		$form.Controls.Add($label)

		#####-- Liste déroulante
		$listbox = New-Object System.Windows.Forms.Listbox
		$listbox.Location = New-Object System.Drawing.Point(10,50)
		$listbox.Size = New-Object System.Drawing.Size(265,150)
		$listbox.BackColor = "Gray"
		$listbox.ForeColor = "White"
		
		if(Test-Path -Path "$generatorDirectory\$filenamePattern*.ps1") {
			# Recherche de la chaine dans tous les fichiers du repertoire avec un format ps1
			Get-Childitem $generatorDirectory | Where-Object {$_.name -like "$filenamePattern*.ps1"} | foreach {
				[void] $listbox.Items.Add($_.name)
			}
		}
		else {
			write-host "`r"
			write-host "--------------------------------------------------------------------------------"
			write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
			write-host "AUCUN FICHIER AVEC LA CHAINE " $filenamePattern " AU FORMAT ps1 DANS LE DOSSIER" $generatorDirectory
			write-host "--------------------------------------------------------------------------------"
			pause
			exit
		}

		$form.Controls.Add($listbox)
		$form.Topmost = $true

		#####-- Affichage de la boite de dialogue
		$choice = $form.ShowDialog()

		if ($choice -eq [System.Windows.Forms.DialogResult]::OK)
		{
			$GENchoice = $listbox.SelectedItems
			write-host "`r"
			write-host "--------------------------------------------------------------------------------"
			write-host "CHOIX DU GENERATEUR CORRECTEMENT EFFECTUE :" $GENchoice
			write-host "--------------------------------------------------------------------------------"
		}
		else
		{
			write-host "`r"
			write-host "--------------------------------------------------------------------------------"
			write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
			write-host "AUCUN TEST N'A ETE SELECTIONNE"
			write-host "--------------------------------------------------------------------------------"
			pause
			exit
		}
		
# CHOIX DU POSTE
		
		#####-- Boite de dialogue
		$form = New-Object System.Windows.Forms.Form
		$form.Text = 'Choix du Poste'
		$form.Size = New-Object System.Drawing.Size(300,300)
		$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
		$form.StartPosition = 'CenterScreen'
		$form.BackColor = "Black"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$form.MaximizeBox = $False

		#####-- Bouton OK
		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Point(10,220)
		$OKButton.Size = New-Object System.Drawing.Size(100,30)
		$OKButton.Text = 'OK'
		$OKButton.BackColor = "Green"
		$OKButton.ForeColor = "White"
		$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$form.AcceptButton = $OKButton
		$form.Controls.Add($OKButton)

		#####-- Bouton Cancel
		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Point(175,220)
		$CancelButton.Size = New-Object System.Drawing.Size(100,30)
		$CancelButton.Text = 'Cancel'
		$CancelButton.BackColor = "Red"
		$CancelButton.ForeColor = "White"
		$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$form.CancelButton = $CancelButton
		$form.Controls.Add($CancelButton)

		#####-- Message
		$label = New-Object System.Windows.Forms.Label
		$label.Location = New-Object System.Drawing.Point(10,10)
		$label.Size = New-Object System.Drawing.Size(265,14)
		$label.Text = "Selectionnez le Poste : "
		$label.ForeColor = "White"
		$form.Controls.Add($label)

		#####-- Liste déroulante
		$listbox = New-Object System.Windows.Forms.Listbox
		$listbox.Location = New-Object System.Drawing.Point(10,30)
		$listbox.Size = New-Object System.Drawing.Size(265,190)
		$listbox.BackColor = "Gray"
		$listbox.ForeColor = "White"
		
		#####-- Eléments de la liste déroulante
		[void] $listbox.Items.Add('22')
		[void] $listbox.Items.Add('24')
		[void] $listbox.Items.Add('25')
		[void] $listbox.Items.Add('26')
		[void] $listbox.Items.Add('27')
		[void] $listbox.Items.Add('75')
		[void] $listbox.Items.Add('81')
		[void] $listbox.Items.Add('83')
		[void] $listbox.Items.Add('85')
		[void] $listbox.Items.Add('87')

		$form.Controls.Add($listbox)
		$form.Topmost = $true

		#####-- Affichage de la boite de dialogue
		$choice = $form.ShowDialog()

		if ($choice -eq [System.Windows.Forms.DialogResult]::OK)
		{
			$Poste = $listbox.SelectedItems
			write-host "`r"
			write-host "--------------------------------------------------------------------------------"
			write-host "CHOIX DU POSTE CORRECTEMENT EFFECTUE : " $Poste
			write-host "--------------------------------------------------------------------------------"
		}
		else
		{
			write-host "`r"
			write-host "--------------------------------------------------------------------------------"
			write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
			write-host "AUCUN POSTE N'A ETE SELECTIONNE"
			write-host "--------------------------------------------------------------------------------"
			pause
			exit
		}
		
# CREATION DU DOSSIER DU POSTE SI NON EXISTENCE
# CHOIX DE LA FE ET DE LA BDD DANS LES REPERTOIRES ASSOCIES AU POSTE SELECTIONNE

		switch ($Poste) {
			('22') {
				[String] $PosteRecord = "05_P22"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('24') {
				[String] $PosteRecord = "06_P24"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('25') {
				[String] $PosteRecord = "07_P25"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('26') {
				[String] $PosteRecord = "08_P26"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('27') {
				[String] $PosteRecord = "09_P27"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('75') {
				[String] $PosteRecord = "10_P75"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('81') {
				[String] $PosteRecord = "11_P81"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('83') {
				[String] $PosteRecord = "12_P83"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}

					}
				}
			}
			('85') {
				[String] $PosteRecord = "13_P85"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
			('87') {
				[String] $PosteRecord = "14_P87"
				[String] $PostePath = Join-Path -Path $RootPath -ChildPath $PosteRecord
				if((Test-Path -Path $PostePath) -eq $FALSE) {
					Copy-Item -Path $XX_PXX -Destination $PostePath -Recurse
					write-host "`r"
					write-host "--------------------------------------------------------------------------------"
					write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
					write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
					write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
					write-host "--------------------------------------------------------------------------------"
					pause
					exit
				}
				else {
					[String] $BddFullPath = Join-Path -Path $PostePath -ChildPath "00_BDD"
					[String] $FEFullPath = Join-Path -Path $PostePath -ChildPath "01_FE"
					if ((Get-Childitem -Path $BddFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE BASE DE DONNEE POSTE DANS LE DOSSIER" $PosteRecord "-> 00_BDD"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					if ((Get-Childitem -Path $FEFullPath) -eq ($NULL)){
						write-host "`r"
						write-host "--------------------------------------------------------------------------------"
						write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
						write-host "METTRE UNE FEUILLE D'ESSAI DANS LE DOSSIER" $PosteRecord "-> 01_FE"
						write-host "--------------------------------------------------------------------------------"
						pause
						exit
					}
					else {
						$FE = @()
						[String] $FEPattern = "FE Poste " + $Poste + "_Cles et CDIV_v*.xlsx"
						$FE = Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern }
						
						if ($FE.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FEUILLE D'ESSAI CONFORME A LA NOMENCLATURE IMPOSEE :" $FEPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA FE
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la FE'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la FE : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $FEFullPath | Where-Object { $_.name -like $FEPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$FEchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA FE :" $FEchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE FE N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						$BDD = @()
						[String] $BDDPattern = "BDD Poste " + $Poste + "_V*.*.xlsm"
						$BDD = Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern }
						
						if ($BDD.Count -eq 0) {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD CONFORME A LA NOMENCLATURE IMPOSEE :" $BDDPattern
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
						
						# CHOIX DE LA BDD
		
						#####-- Boite de dialogue
						$form = New-Object System.Windows.Forms.Form
						$form.Text = 'Choix de la BDD'
						$form.Size = New-Object System.Drawing.Size(300,300)
						$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
						$form.StartPosition = 'CenterScreen'
						$form.BackColor = "Black"
						$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
						$form.MaximizeBox = $False

						#####-- Bouton OK
						$OKButton = New-Object System.Windows.Forms.Button
						$OKButton.Location = New-Object System.Drawing.Point(10,220)
						$OKButton.Size = New-Object System.Drawing.Size(100,30)
						$OKButton.Text = 'OK'
						$OKButton.BackColor = "Green"
						$OKButton.ForeColor = "White"
						$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
						$form.AcceptButton = $OKButton
						$form.Controls.Add($OKButton)

						#####-- Bouton Cancel
						$CancelButton = New-Object System.Windows.Forms.Button
						$CancelButton.Location = New-Object System.Drawing.Point(175,220)
						$CancelButton.Size = New-Object System.Drawing.Size(100,30)
						$CancelButton.Text = 'Cancel'
						$CancelButton.BackColor = "Red"
						$CancelButton.ForeColor = "White"
						$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
						$form.CancelButton = $CancelButton
						$form.Controls.Add($CancelButton)

						#####-- Message
						$label = New-Object System.Windows.Forms.Label
						$label.Location = New-Object System.Drawing.Point(10,10)
						$label.Size = New-Object System.Drawing.Size(265,14)
						$label.Text = "Selectionnez la BDD : "
						$label.ForeColor = "White"
						$form.Controls.Add($label)

						#####-- Liste déroulante
						$listbox = New-Object System.Windows.Forms.Listbox
						$listbox.Location = New-Object System.Drawing.Point(10,30)
						$listbox.Size = New-Object System.Drawing.Size(265,190)
						$listbox.BackColor = "Gray"
						$listbox.ForeColor = "White"
						
						#####-- Eléments de la liste déroulante
						Get-Childitem -Path $BddFullPath | Where-Object { $_.name -like $BDDPattern } | foreach {
							[void] $listbox.Items.Add($_.name)
						}
						
						$form.Controls.Add($listbox)
						$form.Topmost = $true

						#####-- Affichage de la boite de dialogue
						$choice = $form.ShowDialog()
						
						if ($choice -eq [System.Windows.Forms.DialogResult]::OK) {
								$BDDchoice = $listbox.SelectedItems
								write-host "`r"
								write-host "--------------------------------------------------------------------------------"
								write-host "CHOIX DE LA BDD :" $BDDchoice
								write-host "--------------------------------------------------------------------------------"
						}
						else {
							write-host "`r"
							write-host "--------------------------------------------------------------------------------"
							write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
							write-host "AUCUNE BDD N'A ETE SELECTIONNEE"
							write-host "--------------------------------------------------------------------------------"
							pause
							exit
						}
					}
				}
			}
        }


# APPEL DU GENERATEUR DE TEST DE LA FAMILLE POUR LE POSTE CHOISI
		$BddFullPath = "$BddFullPath\$BDDchoice"
		$FEFullPath = "$FEFullPath\$FEchoice"
		[String] $generatorFullPath = "$generatorDirectory\$GENchoice"	
		& $generatorFullPath -BDDFullPath $BddFullPath -FEFullPath $FEFullPath -Poste $Poste -PostePath $PostePath
		
		write-host "`r"
		write-host "--------------------------------------------------------------------------------"
		write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
		write-host "--------------------------------------------------------------------------------"
		
		pause
		exit
}

catch {
		# Gestion des exeptions
		write-host "`r"
		write-host "--------------------------------------------------------------------------------"
		write-host "FIN DU PROCESS DE GENERATION DES SCRIPTS DE TESTS"
		write-host "ERREUR EXCEPTIONNELLE : " $($_.exception.message)
		write-host "--------------------------------------------------------------------------------"
		pause
		exit
}
#####################################################################################################################
# Copyright (c) Siemens S.A.S. 03/2022
# All Rights Reserved, Confidential
# Author		: 	Guillaume POIRIER 
#
# Version 		: 	0.1
# Description 	: 	Manager des generateurs de
#					tests
#
# README 		:	lire la documentation associée
#####################################################################################################################
# Initialisation des parametres d'entree
	PARAM (
		# Identifiant du test generique
		[Parameter(Position=0)]
		[String]
		$IdTestGen,
		# Poste concerne
		[Parameter(Position=1)]
		[String]
		$Poste
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'

		# CHOIX REJEU / ENTIER

		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing

		#####-- Boite de dialogue
		$form = New-Object System.Windows.Forms.Form
		$form.Text = 'Choix de l''execution de la famille de test ' + $IdTestGen + ' pour le poste ' + $Poste
		$form.Size = New-Object System.Drawing.Size(300,125)
		$form.Font = New-Object System.Drawing.Font("Segoe","14",0,2,0)
		$form.StartPosition = 'CenterScreen'
		$form.BackColor = "Black"
		$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
		$form.MaximizeBox = $False

		#####-- Bouton ENTIER
		$AllButton = New-Object System.Windows.Forms.Button
		$AllButton.Location = New-Object System.Drawing.Point(10,45)
		$AllButton.Size = New-Object System.Drawing.Size(100,30)
		$AllButton.Text = 'ENTIER'
		$AllButton.BackColor = "Gray"
		$AllButton.ForeColor = "White"
		$AllButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
		$form.AcceptButton = $AllButton
		$form.Controls.Add($AllButton)

		#####-- Bouton REJEU
		$ReplayButton = New-Object System.Windows.Forms.Button
		$ReplayButton.Location = New-Object System.Drawing.Point(175,45)
		$ReplayButton.Size = New-Object System.Drawing.Size(100,30)
		$ReplayButton.Text = 'REJEU'
		$ReplayButton.BackColor = "Gray"
		$ReplayButton.ForeColor = "White"
		$ReplayButton.DialogResult = [System.Windows.Forms.DialogResult]::No
		$form.CancelButton = $ReplayButton
		$form.Controls.Add($ReplayButton)

		#####-- Message
		$label = New-Object System.Windows.Forms.Label
		$label.Location = New-Object System.Drawing.Point(10,10)
		$label.Size = New-Object System.Drawing.Size(265,30)
		$label.Text = "Quel type d'execution souhaitez-vous?"
		$label.ForeColor = "White"
		$form.Controls.Add($label)
		$form.Topmost = $true

		#####-- Affichage de la boite de dialogue
		$choice = $form.ShowDialog()

		if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
			# Information sur le choix d'execution en mode entier
			$replayTest = 0
			exit $replayTest
		}
		
		if ($choice -eq [System.Windows.Forms.DialogResult]::No) {
			# Information sur le choix d'execution en mode rejeu
			$replayTest = 1
			exit $replayTest
		}
		
		if (($choice -ne [System.Windows.Forms.DialogResult]::Yes) -and ($choice -ne [System.Windows.Forms.DialogResult]::No)) {
			# Gestion du cas ou l'utilisateur choisit d'arreter le programme
			$replayTest = 2
			exit $replayTest
		}
		
}
	catch {
		
		# Gestion des exeptions
		$replayTest = 3
		exit $replayTest
}
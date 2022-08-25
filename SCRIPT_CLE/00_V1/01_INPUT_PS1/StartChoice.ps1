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
		$SpsDirectory
	)
	
	try {
		# Initialisation du paramètre d'execution avec l'option la plus restrictive possible
		# S'il y a une erreur le programme s'arrete
		$ErrorActionPreference = "Stop"
		
		# Initialisation du paramètre d'encodage pour l'ecriture dans le log d'execution du programme
		$PSDefaultParameterValues['Out-File:Encoding'] = 'ascii'

		# CHOIX DU DEBUT

		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing

		#####-- Boite de dialogue
		$form = New-Object System.Windows.Forms.Form
		$form.Text = 'Choix du debut du rejeu'
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
		$label.Text = "Selectionnez le numero du test de depart : "
		$label.ForeColor = "White"
		$form.Controls.Add($label)

		#####-- Liste déroulante
		$listbox = New-Object System.Windows.Forms.Listbox
		$listbox.Location = New-Object System.Drawing.Point(10,50)
		$listbox.Size = New-Object System.Drawing.Size(265,150)
		$listbox.BackColor = "Gray"
		$listbox.ForeColor = "White"
		
		if(Test-Path -Path "$SpsDirectory") {
			# Recherche de la chaine dans tous les fichiers du repertoire avec un format ps1 classe par date de creation croissant
			Get-Childitem "$SpsDirectory" | Sort-Object -Property CreationTime | foreach {
				# Numero prend le nom du fichier
				[String] $number = $_.name
				# Traitement de numero pour ne prendre que le chiffre
				$number = $number.Substring(0,$number.Length - 4)
				# Ajout de numero dans la liste
				[void] $listbox.Items.Add($number)
			}
		}
		else {
			# Gestion de l'erreur impossible de trouver les sps dans l'emplacement
			$Start = -1
			exit $Start
		}

		$form.Controls.Add($listbox)
		$form.Topmost = $true

		#####-- Affichage de la boite de dialogue
		$choice = $form.ShowDialog()

		if ($choice -eq [System.Windows.Forms.DialogResult]::OK)
		{
			# Information sur le choix du test de depart
			$Start = $listbox.SelectedItems
			$Start = $Start.split("_")
			$Start = $Start[$Start.Length - 1]
			exit $Start
		}
		
		else
		{
			# Gestion du cas ou l'utilisateur choisit d'arreter le programme
			$Start = -2
			exit $Start
		}

}
catch {
		
		# Gestion des exeptions
		$Start = -3
		exit $Start
    
}
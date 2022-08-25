//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: confirmSetUp
// Content	: fonction permetant de lancer une commande de confirmation de 
//			  formation
//==============================================================================
int confirmSetUp() {
	
	// Utilitaires
	use system;
	
	// Recuperation de la variable globale "LogExeFullPath"
	string LogExeFullPath = system.getGlobal("LogExeFullPath");
	
	// Ligne de commande pour l'execution du programme .ps1 de confirmation de formation
	string command = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\confirmSettingUp.ps1\"";
	command = command + " -LogExeFullPath \"" + LogExeFullPath + "\"";
	int confirmTest = system.cmd(command);

	return confirmTest;

}
//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: commandLineKey
// Content	: fonction permetant de lancer une commande dans le bandeau mode 
//			  expert
//==============================================================================
int commandLineKey(string key) {
	
	// Utilitaires
	use system;
	
	// Recuperation de la variable globale "LogExeFullPath"
	string LogExeFullPath = system.getGlobal("LogExeFullPath");
	
	// Ligne de commande pour l'execution du programme .ps1 de commande dans le bandeau mode expert
	string command = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\keyAction.ps1\"";
	command = command + " -Key \"" + key + "\"";
	command = command + " -LogExeFullPath \"" + LogExeFullPath + "\"";
	int keyTest = system.cmd(command);

	return keyTest;

}
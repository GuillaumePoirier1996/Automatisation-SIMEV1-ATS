//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: commandLineScreenshot
// Content	: fonction permetant de lancer une commande pour les impressions 
//			  ecran
//==============================================================================
void commandLineScreenshot(string ScreenShotName, string WindowName) {
	
	// Utilitaires
	use system;
	
	// Recuperation des variables globales "LogExeFullPath" et "ScreenShotPath"
	string LogExeFullPath = system.getGlobal("LogExeFullPath");
	string ScreenShotPath = system.getGlobal("ScreenShotPath");
	
	// Ligne de commande pour l'execution du programme .ps1 de commande dans le bandeau mode expert
	string command = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\ScreenShot.ps1\"";
	command = command + " -ScreenShotPath \"" + ScreenShotPath + "\"";
	command = command + " -ScreenShotName \"" + ScreenShotName + "\"";
	command = command + " -appToOpen \"" + WindowName + "\"";
	command = command + " -LogExeFullPath \"" + LogExeFullPath + "\"";
	
	system.cmd(command);

}
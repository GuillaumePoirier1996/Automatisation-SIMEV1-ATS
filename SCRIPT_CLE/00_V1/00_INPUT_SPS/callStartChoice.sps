//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: callStartChoice
// Content	: fonction permetant d'appeler la boite de dialogue "StartChoice"
//==============================================================================
int callStartChoice(string testSpeDirectory) {
	
	// Utilitaires
	use system;
	
	// Ligne de commande pour l'execution de la boite de dialogue "StartChoice"
	string commandStartChoice = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\StartChoice.ps1\"";
	commandStartChoice = commandStartChoice + " -SpsDirectory \"" + testSpeDirectory + "\"";
	int startTest = system.cmd(commandStartChoice);

	return startTest;

}
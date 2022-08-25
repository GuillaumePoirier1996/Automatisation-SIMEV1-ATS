//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: callEndChoice
// Content	: fonction permetant d'appeler la boite de dialogue "EndChoice"
//==============================================================================
int callEndChoice(string testSpeDirectory) {
	
	// Utilitaires
	use system;
	
	// Ligne de commande pour l'execution de la boite de dialogue "EndChoice"
	string commandEndChoice = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\EndChoice.ps1\"";
	commandEndChoice = commandEndChoice + " -SpsDirectory \"" + testSpeDirectory + "\"";
	int endTest = system.cmd(commandEndChoice);

	return endTest;

}

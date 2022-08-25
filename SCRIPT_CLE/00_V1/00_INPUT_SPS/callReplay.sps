//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File	: callReplay
// Content	: fonction permetant d'appeler la boite de dialogue "AskReplay"
//==============================================================================
int callReplay(string idTestGen,string poste) {
	
	// Utilitaires
	use system;
	
	// Ligne de commande pour l'execution de la boite de dialogue "AskReplay"
	string commandReplay = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\AskReplay.ps1\"";
	commandReplay = commandReplay + " -IdTestGen \"" + idTestGen + "\"";
	commandReplay = commandReplay + " -Poste \"" + poste + "\"";
	println(commandReplay);
	int replayTest = system.cmd(commandReplay);

	return replayTest;

}


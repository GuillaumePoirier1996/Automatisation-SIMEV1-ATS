//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: testTraceLogOrganState
// Content	: fonction permetant de tester l'etat d'un organe se trouvant dans
//			  le fichier de traces du CCK avant une action
//==============================================================================
bool testTraceLogOrganStateBeforeAction(string organ,int organState, string h) {
	
	// Utilitaires
	use system;
	
	// Recuperation de la variable globale "LogExeFullPath"
	string LogExeFullPath = system.getGlobal("LogExeFullPath");
	
	// Ligne de commande pour l'execution du programme .ps1 de recherche d'etat d'organe
	//string command = "powershell.exe -ExecutionPolicy Bypass -File \"D:\NEXT_Configuration\Records\SCRIPT_CLE\00_V1\01_INPUT_PS1\ReadOrganStateLogFile_v1.ps1\"";
	string command = "powershell.exe -ExecutionPolicy Bypass -File \"C:\Users\consultant\Desktop\Guillaume\OneDrive - Siemens AG\Documents\Utilitaire_CLE\ReadOrganStateLogFile_v1.ps1\"";
	command = command + " -organ_name \"" + organ + "\" -h \"" + h + "\"";
	command = command + " -LogExeFullPath \"" + LogExeFullPath + "\"";
	int organ_state = system.cmd(command);
	
	// Affichage de tous les parametres
	println("Organe : " + organ + endl + "ETAT : " + organ_state);
	
	// test booleen pour l'ensemble des paramètres
	bool testOrgan = (organ_state == organState);

	return testOrgan;

}
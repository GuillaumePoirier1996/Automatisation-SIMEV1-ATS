//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: stepEnd
// Content	: fonction permetant d'ecrire le resultat d'un step dans le fichier
//			  de resultat
//==============================================================================
void stepEnd(bool exp,variant t,int numStep,string comment) {
	
	// Utilitaires
	use system;
	
	// Affichage du resultat
	t.test(exp,"RESULTAT STEP " + numStep + " : OK (" + comment + ")","RESULTAT STEP " + numStep + " : NOK (" + comment + ")");
	println("----------------------------------------------------------------------------------------------------------------------------------------------------------------" + endl);

}

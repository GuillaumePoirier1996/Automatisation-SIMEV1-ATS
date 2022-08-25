//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: stepPres
// Content	: fonction permetant d'ecrire la presentation d'un step de test
//			  dans le fichier de resultat
//==============================================================================
void stepPres(int numStep,string displayStep) {
	
	// Utilitaires
	use system;
	
	// Affichage du resultat
	println(endl + "----------------------------------------------------------------------------------------------------------------------------------------------------------------");
	println("STEP " + numStep + " : " + displayStep + endl);

}
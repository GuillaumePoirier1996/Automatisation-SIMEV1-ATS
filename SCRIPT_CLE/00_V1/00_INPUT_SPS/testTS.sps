//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: testTS
// Content	: fonction permetant de tester l'etat et la valeur d'une TS
//==============================================================================
bool testTS(variant ts,string tsState,int tsValue) {
	
	// Utilitaires
	use system;
	
	// Recuperation de l'etat de la TS
	string tsStateTest = ts.getState();
	
	// Recuperation de la valeur de la TS
	int tsValueTest = ts.getValue();
	
	// Affichage de tous les parametres
	println("TS : " + ts.getId() + endl + "ETAT COURANT : " + tsStateTest + endl + "VALEUR COURANTE : " + tsValueTest);
	
	// test booleen pour l'ensemble des paramètres
	bool testTs = ((tsStateTest == tsState) && (tsValueTest == tsValue));

	return testTs;

}
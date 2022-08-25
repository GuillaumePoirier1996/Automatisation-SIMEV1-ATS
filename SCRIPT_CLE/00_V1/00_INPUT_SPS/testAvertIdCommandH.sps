//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: testAvertIdCommandH
// Content	: fonction permetant de tester l'id, l'id de la commande et 
//			  l'horodatage d'un avertissement
//==============================================================================
bool testAvertIdCommandH(string h,string idAvert,string idCommand) {
	
	// Utilitaires
	use system;
	use S_CCK_MCMD_AVERTISSEMENT;
	
	// Recuperation de l'alarme au format json
	json avert = S_CCK_MCMD_AVERTISSEMENT.getState();

	// Recuperation de l'id
	json id_avert = avert.getValueByKey("ID_AVERTISSEMENT");
	string id_avert1 = id_avert.getValueAsString(true);

	// Recuperation de l'argument
	json id_command = avert.getValueByKey("ID_COMMANDE");
	string id_command1 = id_command.getValueAsString(true);
	
	// Recuperation de l'horodatage
	date hav = S_CCK_MCMD_AVERTISSEMENT.getLastChange();
	
	// Affichage de tous les parametres
	println("ID AVERTISSEMENT : " + id_avert1 + endl + "ID COMMANDE : " + id_command1 + endl + "HORODATAGE AVERTISSEMENT : " + hav);
	
	// transformation de h en format date
	date h1 = h;
	
	// test booleen pour l'ensemble des paramètres
	bool testAvert = ((id_command1 == idCommand) && (id_avert1 == idAvert) && (hav >= h1));

	return testAvert;

}
//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: testAlarmIdArgH
// Content	: fonction permetant de tester l'id et l'horodatage
//			  d'une alarme
//==============================================================================
bool testAlarmIdH(string h,string idAlarm){
	
	// Utilitaires
	use system;
	use S_CCK_ALARME;
	
	// Recuperation de l'alarme au format json
	json alarm = S_CCK_ALARME.getState();

	// Recuperation de l'id
	json id_alarm = alarm.getValueByKey("ID_ALARME");
	string id_alarm1 = id_alarm.getValueAsString(true);
	
	// Recuperation de l'horodatage
	date hal = S_CCK_ALARME.getLastChange();
	
	// Affichage de tous les paramètres
	println("ID ALARME : " + id_alarm1 + endl + "HORODATAGE ALARME : " + hal);
	
	// transformation de h en format date
	date h1 = h;
	
	// test booleen pour l'ensemble des paramètres
	bool testAlarm = ((id_alarm1 == idAlarm) && (hal >= h1));
	
	return testAlarm;
	
}
//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File		: testAlarmIdArgH
// Content	: fonction permetant de tester l'id, l'argument et l'horodatage
//			  d'une alarme
//==============================================================================
bool testAlarmIdArgH(string h,string idAlarm,string Arg) {
	
	// Utilitaires
	use system;
	use S_CCK_ALARME;
	
	// Recuperation de l'alarme au format json
	json alarm = S_CCK_ALARME.getState();

	// Recuperation de l'id
	json id_alarm = alarm.getValueByKey("ID_ALARME");
	string id_alarm1 = id_alarm.getValueAsString(true);

	// Recuperation de l'argument
	json arg_alarm = alarm.getValueByKey("ARGUMENTS");
	string arg_alarm1 = arg_alarm.getValueAsString(true);
	
	// Recuperation de l'horodatage
	date hal = S_CCK_ALARME.getLastChange();
	
	// Affichage de tous les parametres
	println("ID ALARME : " + id_alarm1 + endl + "ARGUMENT : " + arg_alarm1 + endl + "HORODATAGE ALARME : " + hal);
	
	// transformation de h en format date
	date h1 = h;
	
	// test booleen pour l'ensemble des paramètres
	bool testAlarm = ((id_alarm1 == idAlarm) && (hal >= h1) && (Arg == arg_alarm1));

	return testAlarm;

}
//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File     : CTRLAllFnit
// Content  : Booleen indiquant si tous les FNIT / FNIT PRS / FNAU du domaine ATS sont
//			  bien à l'état inactif
//==============================================================================

bool CTRLAllFnitFnau (int poste){
	
	use system;
	
	// Recherche dans le fichier de données qui référence toutes les TS
	string file_name = "D:\NEXT\Data\Csv\NEXT_ts.csv";

	// Vérifie si le fichier .csv existe
	if (system.existCSV(file_name))
	{
		
		string separator = ";";
		bool areLinesIndexed = false;
		bool areColumnsIndexed = true;

		// Charge le contenu du fichier dans une table
		if (system.loadCSV(file_name, separator, areLinesIndexed, areColumnsIndexed))
		{
			// Recherche à partir de la première ligne de la colonne "ID"
			int line_index = 0;
			string column_name = "ID";
			
			// Initialisation du nombre de ts testé à zéro
			int nb_ts = 0;
			
			string ts = system.getTableValue(file_name, line_index , column_name);
			
			while(ts != "") {
				
				// Expressions recherchées dans le fichier .csv
				// TS du FNAU
				string reg_exp1 = "S_CCK_MFCO_FNAU_" + poste + "\w+_FNAU_ASPECT_CTRL";
				// TS du FNIT
				string reg_exp2 = "S_CCK_MFCO_SIG_" + poste + "\w+_FNIT_ASPECT_CTRL";
				// TS du FNIT PRS
				string reg_exp3 = "S_CCK_MFCO_SIG_" + poste + "\w+_FNIT_PRS_ASPECT_CTRL";
				
				// TEST DE TOUS LES FNAU
				if(system.match(ts, reg_exp1)) {
					
					variant tsValue = get(ts);
					
					if(tsValue.getValue() != 1) {
						// Prise de l'identifiant du FNAU
						for(int i = 16 ; ts.substring(i,i) != "_" ; i++){
							string id_FnitFnau = ts.substring(16,i);
						}
						
						println("La ts " + ts + " indique que le FNAU " + id_FnitFnau + " est a l'etat " + tsValue.getState());
						return false;
					}
					
					nb_ts++;
				}
				
				// TEST DE TOUS LES FNIT
				if(system.match(ts, reg_exp2)) {
					
					variant tsValue = get(ts);
					
					if(tsValue.getValue() != 1) {
						// Prise de l'identifiant du FNIT
						for(int i = 15 ; ts.substring(i,i) != "_" ; i++){
							string id_FnitFnau = ts.substring(15,i);
						}
						
						println("La ts " + ts + " indique que le FNIT " + id_FnitFnau + " est a l'etat " + tsValue.getState());
						return false;
					}
					
					nb_ts++;
				}
				
				// TEST DE TOUS LES FNIT PRS
				if(system.match(ts, reg_exp3)) {
					
					variant tsValue = get(ts);
					
					if(tsValue.getValue() != 1) {
						// Prise de l'identifiant du FNIT PRS
						for(int i = 15 ; ts.substring(i,i) != "_" ; i++){
							string id_FnitFnau = ts.substring(15,i);
						}
						
						println("La ts " + ts + " indique que le FNIT PRS " + id_FnitFnau + " est a l'etat " + tsValue.getState());
						return false;
					}
					
					nb_ts++;
				}

				line_index++;
				string ts = system.getTableValue(file_name, line_index , column_name);
			}
			println("Toutes les Ts des FNIT / FNIT PRS / FNAU controlees (" + nb_ts + " au total) sont a l'etat inactif");
			return true;
		}
		else
		{
			println("Impossible de charger le contenu du fichier " + file_name + " dans un tableau");
			return false;
		}
		
	}
	else
	{
		println("Le fichier " + file_name + " n'existe pas");
		return false;
	}

}

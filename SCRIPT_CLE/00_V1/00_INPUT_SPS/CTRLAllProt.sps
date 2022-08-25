//==============================================================================
// Copyright (c) Siemens S.A.S. 2022
// All Rights Reserved, Confidential
// Author	: Guillaume POIRIER
//
// Version	: 1.0
// File     : CTRLAllProt
// Content  : Booleen indiquant si toutes les protection du domaine ATS sont
//			  bien à l'état inactif
//==============================================================================

bool CTRLAllProt (){
	
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
				// TS de l'identifiant de ZEP
				string reg_exp1 = "S_CCK_MFCO_ZEP_\w+_ZEP_LABEL_ASPECT_CTRL";
				// TS de l'identifiant de SEL
				string reg_exp2 = "S_CCK_MFCO_SEL_\w+_SEL_LABEL_ASPECT_CTRL";
				// TS du trait de ZEP
				string reg_exp3 = "S_CCK_MFCO_ZEP_\w+_ZEP_ASPECT_CTRL";
				// TS du trait de SEL
				string reg_exp4 = "S_CCK_MFCO_SEL_\w+_SEL_ASPECT_CTRL";
				
				if(system.match(ts, reg_exp1) || system.match(ts, reg_exp2) || system.match(ts, reg_exp3) || system.match(ts, reg_exp4)) {
					
					variant tsValue = get(ts);
					
					if(tsValue.getValue() != 1) {
						// Prise de l'identifiant de la protection
						for(int i = 15 ; ts.substring(i,i) != "_" ; i++){
							string id_prot = ts.substring(15,i);
						}

						println("La ts " + ts + " indique que la " + ts.substring(11,13) + " " + id_prot + " est a l'etat " + tsValue.getState());
						return false;
					}
					
					nb_ts++;
				}

				line_index++;
				string ts = system.getTableValue(file_name, line_index , column_name);
			}
			println("Toutes les Ts des protections controlees (" + nb_ts + " au total) sont a l'etat inactif");
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

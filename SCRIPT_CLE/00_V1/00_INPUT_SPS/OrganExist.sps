//==============================================================================
// Copyright (c) Siemens S.A.S. 01/2022
// All Rights Reserved, Confidential
// Author      : S DIALLO

// File        : OrganExist.sps
// Description :
//		1) Vérifie si l'identifiant d'un organe existe dans la liste des organes
//		2) Charge le csv NEXT_tsOrgan.csv
//==============================================================================


bool OrganExist(string organ_id)
{
	use system;
	bool result = false;
	
	string file = "D:\NEXTTS\Data\Csv\NEXT_tsOrgan.csv";
	if (system.existCSV(file))
	{
		// Charge le fichier CSV et retourne la liste de toutes les valeurs de la colonne (ID)
		string column_name = "ID";
		string separator = ";";
		variant array_ID = system.loadCsvAndGetAllValuesByColumnName(file, column_name, separator);
		
		// Taille du tableau
		int array_size = array_ID.size();
		
		// Vérifie si l'iditifiant de l'organe figure dans la liste
		int i = 0;
		bool break = true;
		string array_value;
		while (i < array_size && break)
		{
			array_value = array_ID[i];
			if (array_value == organ_id)
			{
				result = true;
				break = false;
			}
			i = i + 1;
		}
	}
	
	return result;
}
void test(){
	use system;
	
	date h = h.now();
	string h1 = h;
	println(h1);
	SetIncident("83AUATR8335","CONTROL_STATE_TO_ACTIVE","ACTIF");
	bool act = testTraceLogOrganState("83AUATR8335",1,h1);
	
	delay(3);

	date h = h.now();
	string h1 = h;
	SetIncident("83AUATR8335","CONTROL_STATE_TO_INACTIVE","INACTIF");
	bool inact = testTraceLogOrganState("83AUATR8335",0,h1);
	

}

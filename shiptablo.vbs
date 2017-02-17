Function dai1(Arg1, psFrom, texto, resFile)
	posSt1=InStr(psFrom, Arg1, "title="+chr(34)+texto)
	posSt=InStr(posSt1, Arg1, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, Arg1, "<")
	data=Replace(Replace(Mid(Arg1, posSt, posEnd-posSt),",",""), "&#179;","")
	'data=Replace(data, " km","")
	
	if(texto="base time to warp") then
	   data=Replace(data, " s","")
	end if

	if(texto="max. targeting range") then
	   data=Replace(data, " km","")
	end if

	if ((texto="drone capacity")or(texto="cargo capacity")or(texto="ship signature radius")) then
	   data=Replace(data, " m","")
	end if
	
	resFile.write (data+",")	
	dai1=posEnd
End Function	

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("shipz.csv",True)  
resFile.write ("ShipName,Nation,Type,TrainingTime,Powergrid,CPU,Capacitor,high slots,launcher,turret,medium slots,low slots,rigs,calibration," + _
"max. velocity,Inertia Modifier,Warp Speed,Base Time to Warp,drone capacity,drone bandwith,max. targeting range,max.locked tagets,ship signature radius,scan res," + _
"structure hitpoints,cargo capacity,shields,armor,")
resFile.write (vbCrLf)

Dim shipz
shipz = Split ("Impairor Ibis Velator Reaper Executioner Inquisitor Tormentor Crucifier Punisher Magnate Condor Kestrel Griffin Merlin Heron "+ _
"Atron Navitas Tristan Maulus Incursus Imicus Slasher Burst Breacher Vigil Rifter Probe Imperial_Navy_Slicer Crucifier_Navy_Issue "+ _
"Caldari_Navy_Hookbill Griffin_Navy_Issue Federation_Navy_Comet Maulus_Navy_Issue Republic_Fleet_Firetail Vigil_Fleet_Issue Vengeance "+ _
"Retribution Hawk Harpy Ishkur Enyo Wolf Jaguar Anathema Purifier Manticore Buzzard Nemesis Helios Cheetah "+ _
"Hound Sentinel 	Kitsune 	Keres 	Hyena 	Crusader 	Malediction 	Crow 	Raptor 	Taranis 	Ares 	Claw "+ _
"Stiletto Deacon 	Kirin 	Thalia 	Scalpel 	Coercer 	Dragoon 	Cormorant 	Corax 	Catalyst 	Algos 	Thrasher "+ _
"Talwar 	Pontifex 	Stork 	Magus 	Bifrost 	Heretic 	Flycatcher 	Eris 	Sabre 	Confessor 	Jackdaw 	Hecate "+ _
"Svipul 	Arbitrator 	Maller 	Augoror 	Omen 	Moa 	Blackbird 	Osprey 	Caracal 	Celestis 	Exequror 	Vexor "+ _
"Thorax 	Stabber 	Rupture 	Bellicose 	Scythe 	Augoror_Navy_Issue 	Omen_Navy_Issue 	Caracal_Navy_Issue 	Osprey_Navy_Issue "+ _
"Exequror_Navy_Issue 	Vexor_Navy_Issue 	Scythe_Fleet_Issue 	Stabber_Fleet_Issue 	Zealot 	Sacrilege 	Eagle 	Cerberus "+ _
"Ishtar 	Deimos 	Vagabond 	Muninn 	Devoter 	Onyx 	Phobos 	Broadsword 	Guardian 	Basilisk 	Oneiros 	Scimitar "+ _
"Curse 	Pilgrim 	Falcon 	Rook 	Arazu 	Lachesis 	Huginn 	Rapier 	Legion 	Tengu 	Proteus 	Loki 	Prophecy "+ _
"Harbinger 	Oracle 	Drake 	Ferox 	Naga 	Brutix 	Myrmidon 	Talos 	Hurricane 	Cyclone 	Tornado 	Harbinger_Navy_Issue "+ _
"Drake_Navy_Issue 	Brutix_Navy_Issue 	Hurricane_Fleet_Issue 	Damnation 	Absolution 	Vulture 	Nighthawk 	Astarte 	Eos "+ _
"Claymore 	Sleipnir 	Abaddon 	Apocalypse 	Armageddon 	Scorpion 	Raven 	Rokh 	Hyperion 	Megathron 	Dominix 	Tempest "+ _
"Typhoon 	Maelstrom 	Apocalypse_Navy_Issue 	Armageddon_Navy_Issue 	Scorpion_Navy_Issue 	Raven_Navy_Issue "+ _
"Megathron_Navy_Issue 	Dominix_Navy_Issue 	Tempest_Fleet_Issue 	Typhoon_Fleet_Issue 	Redeemer 	Widow 	Sin "+ _
"Panther 	Paladin 	Golem 	Kronos_(Ship) 	Vargur 	Archon 	Chimera 	Thanatos 	Nidhoggur 	Aeon 	Wyvern 	Nyx "+ _
"Hel 	Revelation 	Phoenix 	Moros 	Naglfar 	Apostle 	Minokawa 	Ninazu 	Lif 	Avatar 	Leviathan 	Erebus "+ _
"Ragnarok 	Providence 	Charon 	Obelisk 	Fenrir 	Ark 	Rhea 	Anshar 	Nomad 	Bestower 	Sigil 	Badger 	Tayra 	Nereus "+ _
"Kryos 	Epithal 	Miasmos 	Iteron_Mark_V 	Hoarder 	Mammoth 	Wreathe 	Prorator 	Impel 	Crane 	Bustard 	Viator "+ _
"Occator 	Mastodon 	Prowler 	Venture 	Prospect 	Endurance 	Covetor 	Retriever 	Procurer 	Hulk 	Skiff "+ _
"Mackinaw 	Noctis 	Orca 	Bowhead 	Rorqual 	Dramiel 	Cruor 	Worm 	Garmur 	Succubus 	Daredevil 	Astero 	Cynabal "+ _
"Ashimmu 	Gila 	Orthrus 	Phantasm 	Vigilant 	Stratios 	Machariel 	Bhaalgorn 	Rattlesnake 	Barghest 	Nightmare "+ _
"Vindicator Nestor Vehement Revenant Vendetta 	Vanquisher 	Apotheosis 	Interbus_Shuttle 	Leopard 	Echelon 	Echo "+ _
"Hematos 	Immolator 	Taipan 	Violator 	Zephyr 	Gold_Magnate 	Inner_Zone_Shipping_Imicus 	Sarum_Magnate 	Silver_Magnate "+ _
"Sukuuvestaa_Heron 	Tash-Murkon_Magnate 	Vherokior_Probe 	Cambion 	Freki 	Malice 	Utu 	Chremoas 	Imp 	Whiptail "+ _
"Aliastra_Catalyst 	Inner_Zone_Shipping_Catalyst 	Intaki_Syndicate_Catalyst 	InterBus_Catalyst 	Nefantar_Thrasher "+ _
"Quafe_Catalyst 	Guardian-Vexor 	Victorieux_Luxury_Yacht 	Adrestia 	Mimir 	Vangel 	Fiend 	Etana 	Chameleon 	Moracha "+ _
"Gnosis 	Scorpion_Ishukone_Watch 	Apocalypse_Imperial_Issue 	Armageddon_Imperial_Issue 	Raven_State_Issue 	Megathron_Federate_Issue "+ _
"Tempest_Tribal_Issue Primae Miasmos_Amastris_Edition Miasmos_Quafe_Ultra_Edition Miasmos_Quafe_Ultramarine_Edition" )     

for j = 0 to UBound(shipz)
	nazwa=shipz(j)

	xmlhttp.open "get", "http://wiki.eveuniversity.org/"+nazwa, false
	xmlhttp.send
	MyText= xmlhttp.responseText

	'<div class="shipname">Bantam</div>
	posShipSt=InStr(1, MyText,"div class="+chr(34)+"shipname")+21
	posShipEnd=InStr(posShipSt, MyText, "<")
	'MsgBox " "+cstr(posShipSt)+" "+cstr(posShipEnd)
	shipName=Mid(MyText, posShipSt, posShipEnd-posShipSt)
	resFile.write (shipName+",")

	
	if nazwa="Vehement"  then            'sabaki!!!
       resFile.write ("Serpentis,Pirate Faction Ships,")
    else 	   
	   '<td class="faction">Caldari State</td>
	   posNatSt=InStr(1, MyText,"td class="+chr(34)+"faction")+19
	   posNatEnd=InStr(posNatSt, MyText, "<")
	   shipNat=Mid(MyText, posNatSt, posNatEnd-posNatSt)
	   resFile.write (shipNat+",")
    	

       'title="Category:Ship Database">Ship Database</a></li><li><a href="/Category:Standard_Frigates" title="Category:Standard Frigates">Standard Frigates</a>
	   posCateSt1=InStr(1, MyText,"Ship Database</a></li><li><a href="+chr(34)+"/Category:")
	   posCateSt=InStr(posCateSt1, MyText,"title="+chr(34)+"Category:")+16
	   posCateEnd=InStr(posCateSt, MyText, ">")-1
	   shipCateg=Mid(MyText, posCateSt, posCateEnd-posCateSt)
	   resFile.write (shipCateg+",")
	end if
    
	posTrainingSt1=InStr(posShipEnd, MyText, "Training Time")
	posTrainingSt=InStr(posTrainingSt1, MyText, "div class="+chr(34)+"value")+18
	posTrainingEnd=InStr(posTrainingSt, MyText, "<")
	trainingTime=Mid(MyText, posTrainingSt, posTrainingEnd-posTrainingSt)
	resFile.write (trainingTime+",")

	kanec=dai1(MyText, posTrainingEnd, "powergrid", resFile)
	kanec=dai1(MyText, kanec, "cpu output", resFile)
	kanec=dai1(MyText, kanec, "capacitor", resFile)
	kanec=dai1(MyText, kanec, "high slots", resFile)
	kanec=dai1(MyText, kanec, "launcher slots", resFile)
	kanec=dai1(MyText, kanec, "turret slots", resFile)
	kanec=dai1(MyText, kanec, "middle slots", resFile)
	kanec=dai1(MyText, kanec, "low slots", resFile)
	kanec=dai1(MyText, kanec, "rigs", resFile)
	kanec=dai1(MyText, kanec, "calibration", resFile)
	kanec=dai1(MyText, kanec, "max. velocity", resFile)
	kanec=dai1(MyText, kanec, "inertia modifier", resFile)

	' BUG IN THEIR CODE - must be Warp Speed
	kanec=dai1(MyText, kanec, "inertia modifier", resFile)

	kanec=dai1(MyText, kanec, "base time to warp", resFile)
	kanec=dai1(MyText, kanec, "drone capacity", resFile)
	kanec=dai1(MyText, kanec, "drone bandwith", resFile)
	kanec=dai1(MyText, kanec, "max. targeting range", resFile)
	kanec=dai1(MyText, kanec, "max. locked targets", resFile)
	'kanec=dai1(MyText, kanec, "Gravimetric sensor strength", resFile)
	kanec=dai1(MyText, kanec, "ship signature radius", resFile)
	kanec=dai1(MyText, kanec, "scan resolution", resFile)
	
	kanec=dai1(MyText, kanec, "structure hitpoints", resFile)
	kanec=dai1(MyText, kanec, "cargo capacity", resFile)
	kanec=dai1(MyText, kanec, "shields", resFile)
	kanec=dai1(MyText, kanec, "armor", resFile)
	
	
	resFile.write (vbCrLf)
	 WScript.Sleep 1500
next

resFile.Close


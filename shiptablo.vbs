Function dai1(Arg1, psFrom, texto, resFile)
	posSt1=InStr(psFrom, Arg1, "title="+chr(34)+texto)
	posSt=InStr(posSt1, Arg1, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, Arg1, "<")
	data=Replace(Replace(Mid(Arg1, posSt, posEnd-posSt),",",""), "&#179;","")
	resFile.write (data+",")	
	dai1=posEnd
End Function	

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("shipz.csv",True)  
resFile.write ("ShipName,TrainingTime,Powergrid,CPU,Capacitor,high slots, launcher, turret, medium slots, low slots, rigs, calibration," + _
"max. velocity,Inertia Modifier,Warp Speed,Base Time to Warp,max. targeting range, max.locked tagets, signature radius, scan res," + _
"structure hitpoints, cargo capacity, shields, armor,")
resFile.write (vbCrLf)

Dim shipz
shipz = Split ("Impairor Kestrel Griffin Merlin Heron Atron Navitas Tristan Maulus Incursus Imicus Slasher Burst Breacher")     

for j = 0 to UBound(shipz)
	nazwa=shipz(j)

	xmlhttp.open "get", "http://wiki.eveuniversity.org/"+nazwa, false
	xmlhttp.send
	MyText= xmlhttp.responseText

	posShipSt=InStr(1, MyText,"div class="+chr(34)+"shipname")+21
	posShipEnd=InStr(posShipSt, MyText, "<")
	'MsgBox " "+cstr(posShipSt)+" "+cstr(posShipEnd)
	shipName=Mid(MyText, posShipSt, posShipEnd-posShipSt)
	resFile.write (shipName+",")
	
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


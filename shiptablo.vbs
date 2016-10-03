
	
Function dai1(Arg1, psFrom, texto, resFile)
	posSt1=InStr(psFrom, Arg1, "title="+chr(34)+texto)
	posSt=InStr(posSt1, Arg1, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, Arg1, "<")
	data=Replace(Mid(Arg1, posSt, posEnd-posSt),",","")
	resFile.write (data+",")	
	dai1=posEnd
End Function	

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("shipz.csv",True)  
resFile.write ("ShipName,TrainingTime,Powergrid,CPU,Capacitor,high slots, launcher, turret, medium slots, low slots, rigs, calibration," + _
"max. velocity,Inertia Modifier,Warp Speed,Base Time to Warp,max. targeting range, max.locked tagets, signature radius, scan res")
resFile.write (vbCrLf)

Dim shipz
shipz = Split ("Impairor Kestrel 	Griffin 	Merlin 	Heron 	Atron 	Navitas 	Tristan 	Maulus 	Incursus 	Imicus 	Slasher 	Burst 	Breacher")     

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

	posEnd=posTrainingEnd

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"powergrid")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"cpu output")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"capacitor")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")	

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"high slots")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		
	
	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"launcher slots")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		
	
	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"turret slots")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		
	
	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"middle slots")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		
	
	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"low slots")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"rigs")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"calibration")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"max. velocity")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		

	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"inertia modifier")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		
	
	' BUG IN THEIR CODE - must be Warp Speed
	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"inertia modifier")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")		
	
	posSt1=InStr(posEnd, MyText, "title="+chr(34)+"base time to warp")
	posSt=InStr(posSt1, MyText, "div class="+chr(34)+"item att-value")+27
	posEnd=InStr(posSt, MyText, "<")
	data=Replace(Mid(MyText, posSt, posEnd-posSt),",","")
	resFile.write (data+",")	
	
	kanec=dai1(MyText, posEnd, "max. targeting range", resFile)
	kanec=dai1(MyText, kanec, "max. locked targets", resFile)
	'kanec=dai1(MyText, kanec, "Gravimetric sensor strength", resFile)
	kanec=dai1(MyText, kanec, "ship signature radius", resFile)
	kanec=dai1(MyText, kanec, "scan resolution", resFile)
	
	
	resFile.write (vbCrLf)
	 WScript.Sleep 1500
next

resFile.Close


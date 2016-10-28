
set xmlhttp = createobject ("msxml2.xmlhttp.3.0")

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("bluprin.csv",True)  
resFile.write ("ItemId,Blueprint,Price,Zatrata,ProfitMon,ProfitProc")
resFile.write (vbCrLf)

'<strong>Mjolnir Light Missile</strong>
'

'find third  <tr class="catheader">
'after it find second <td   --- name
'after it <span style="color  --- current price
'next  <span style="color ---- zatrata
'next  <span style="color ---- profit

for j = 800 to 830

	xmlhttp.open "get", "http://eve-industry.org/calc/getInfo/?id="+CStr(j)+"&me=10&te=20&runs=1&enc=5&dc1=5&dc2=5&decryptor=0&techlevel=1&byoc=false&buildcopy=false&cme=10&cte=20&advanced_industry=5&materials_modifier=1&c_materials_modifier=1&skill_m=0.80&skill_mc=0.80&skill_te=0.75&skill_me=0.75&skill_c=0.75&skill_i=1.0&implant_m=1.0&implant_mc=1.0&implant_te=1.0&implant_me=1.0&implant_c=1.0&implant_i=1.0&facility_m=1&facility_mc=1&facility_te=1&facility_me=1&facility_c=1&facility_i=1&solarSystem_m=Osmon&solarSystem_mc=Osmon&solarSystem_te=Osmon&solarSystem_me=Osmon&solarSystem_c=Osmon&solarSystem_i=Osmon&taxRate_m=10&taxRate_mc=10&taxRate_te=10&taxRate_me=10&taxRate_c=10&taxRate_i=10&pi=false", false
	xmlhttp.send
	MyText= xmlhttp.responseText

	if len(MyText)>0 then

	    resFile.write (CStr(j)+",")
	
		nazva1St=InStr(1, MyText,"tr class="+chr(34)+"catheader")+1
		nazva2St=InStr(nazva1St, MyText,"tr class="+chr(34)+"catheader")+1
		nazva3St=InStr(nazva2St, MyText,"tr class="+chr(34)+"catheader")+1
		nazva4St=InStr(nazva3St, MyText,"tr class="+chr(34)+"catheader")+1
		nazvatd1St=InStr(nazva4St, MyText,"<td")+1
		nazvatd2St=InStr(nazvatd1St, MyText,"<td")+4
			
		nazvaEnd=InStr(nazvatd2St, MyText, "</td")

		Name=Mid(MyText, nazvatd2St, nazvaEnd-nazvatd2St)
		'MsgBox Name
		resFile.write (Name+",")

		currPr1St=InStr(nazvaEnd, MyText,"<span style="+chr(34)+"color")+1
		currPrSt=InStr(currPr1St, MyText,">")+1
		currPrEnd=InStr(currPrSt, MyText, "ISK")
		Price=Mid(MyText, currPrSt, currPrEnd-currPrSt)
		PriceNum=CDbl(  Replace( Replace(Price,",","")," ","") )
		resFile.write (CStr(PriceNum)+",")
		
		zatrata1St=InStr(currPrEnd, MyText,"<span style="+chr(34)+"color")+1
		zatrataSt=InStr(zatrata1St, MyText,">")+1
		zatrataEnd=InStr(zatrataSt, MyText, "ISK")
		Zatrata=Mid(MyText, zatrataSt, zatrataEnd-zatrataSt)
		ZatrataNum=CDbl(  Replace(Replace( Replace(Zatrata,",","")," ",""),"-","") )
		resFile.write (CStr(ZatrataNum)+",")
		
		ProfitMon=PriceNum-ZatrataNum
		resFile.write (CStr(ProfitMon)+",")

		ProfitProc=100*ProfitMon/PriceNum
		resFile.write (CStr(ProfitProc)+",")
		
		resFile.write (vbCrLf)
	End If	
	WScript.Sleep 1500
		 
next

resFile.Close


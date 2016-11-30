URL="https://eve-central.com/home/tradefind_display.html?qtype=SystemToRegion&newsearch=1&fromt=30000142&to=10000002"

htmlName="from4mis.html"

URL1="https://eve-central.com/home/tradefind_display.html?set=1&fromt="
URL2="&to="
KUDA="30000142"  'Jita
URL3="&qtype=Systems&age=24&minprofit=10000&size=1000000&limit=2000&sort=sprofit&prefer_sec=1"


set xmlhttp = createobject ("msxml2.xmlhttp.3.0")


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile(htmlName,True)  

resFile.write ("<html><head><script src=sorttable.js></script></head><body><table  border=1 class=sortable>" & vbCrLf)
resFile.write ("<tr> <th>from</th> <th>to</th> <th>jumps</th> <th>what</th>  <th>sell</th>" )
resFile.write ("<th>buy</th> <th>count</th><th>polnoe</th> <th>tot money</th> <th>tot prof</th> <th>perc prof</th>  </tr> " & vbCrLf)



Dim stankas
stankas= Split ("30005216 30005215 30005214 30005202 30005201 30005199 30005198 30005015 30002761 30002764 30002765 30002768 30002800 30002801 30002802 30000139 30000144 30000142")


for j = 0 to UBound(stankas)
	stanka=stankas(j)

	URL=URL1+stanka+URL2+KUDA+URL3
	'MsgBox URL
	xmlhttp.open "get", URL, false
	xmlhttp.send
	MyText= xmlhttp.responseText

	startpos=1

	do while startpos>0
	    fromPosSt=InStr(startpos, MyText,"<b>From:</b>")
		if fromPosSt=0 then exit do
	    fromPosEnd=InStr(fromPosSt, MyText,"</td>")
	    fromStr=Mid(MyText, fromPosSt+12, fromPosEnd-fromPosSt-12)

	    toPosSt=InStr(fromPosEnd, MyText,"<b>To:</b>")
	    toPosEnd=InStr(toPosSt, MyText,"</td>")
	    toStr=Mid(MyText, toPosSt+10, toPosEnd-toPosSt-10)

	    jumpsPosSt=InStr(toPosEnd, MyText,"<b>Jumps:</b>")
	    jumpsPosEnd=InStr(jumpsPosSt, MyText,"</td>")
	    jumpsStr=Mid(MyText, jumpsPosSt+13, jumpsPosEnd-jumpsPosSt-13)
		
	    itemPosSt=InStr(jumpsPosEnd, MyText,"<b>Type:</b>")
	    itemPosEnd=InStr(itemPosSt, MyText,"</td>")
	    itemStr=Mid(MyText, itemPosSt+12, itemPosEnd-itemPosSt-12)
		'and link to get volume!!
		itemStr=Replace(itemStr, "quicklook",  "https://eve-central.com/home/quicklook" )

	    sellprPosSt=InStr(itemPosEnd, MyText,"<b>Selling:</b>")
	    sellprPosEnd=InStr(sellprPosSt, MyText,"ISK")
	    sellprStr=Mid(MyText, sellprPosSt+15, sellprPosEnd-sellprPosSt-15)

	    buyprPosSt=InStr(sellprPosEnd, MyText,"<b>Buying:</b>")
	    buyprPosEnd=InStr(buyprPosSt, MyText,"ISK")
	    buyprStr=Mid(MyText, buyprPosSt+14, buyprPosEnd-buyprPosSt-14)

		unitsPosSt=InStr(buyprPosEnd, MyText,"<b>Units tradeable:</b>")
	    unitsPosEnd=InStr(unitsPosSt, MyText,"(")
	    unitsStr=Trim(Mid(MyText, unitsPosSt+23, unitsPosEnd-unitsPosSt-23))
		
	    units1PosEnd=InStr(unitsPosSt, MyText,")")
	    unitsPolnStr=Trim(Mid(MyText, unitsPosSt+23, units1PosEnd-unitsPosSt-23))

	    profitPosSt=InStr(unitsPosEnd, MyText,"<i>Profit per trip:</i>")
	    profitPosEnd=InStr(profitPosSt, MyText,"ISK")
	    profitStr=Trim(Mid(MyText, profitPosSt+28, profitPosEnd-profitPosSt-28))

		' count 
		totalmoneyNum=CDbl(Replace(sellprStr,",",""))*CLng(Replace(unitsStr,",",""))
		totalmoneyStr=CStr(totalmoneyNum)
		
		profit1Num=CDbl(Replace(profitStr,",",""))
		profit1Str=CStr(profit1Num)

		
		profitpercentnum=CStr( profit1Num/totalmoneyNum	)
		
		if (profit1Num/totalmoneyNum)>CDbl(0.05) and profit1Num>500000 then 
		resFile.write ("<tr> <td>"&fromStr &"</td> <td>"&toStr &"</td> <td>"&jumpsStr &"</td> <td>"&itemStr &"</td>  <td>"&sellprStr &"</td>" )
        resFile.write ("<td>"&buyprStr&"</td> <td>"&unitsStr&"</td><td>"&unitsPolnStr&"</td><td>"&FormatCurrency(totalmoneyStr)&"</td> <td>"&FormatCurrency(profitStr)&"</td> <td>"&FormatPercent(profit1Num/totalmoneyNum)&"</td>  </tr> " & vbCrLf)
		end if
		
		
		startpos=InStr(profitPosEnd, MyText,"<b>From:</b>")
    loop 
	
next

resFile.write ("</table></body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start " + htmlName



'Apanake 	0.5		30005216
'Avyuh 	0.6		30005215
'Ashokon 	0.7		30005214
'Emsar 	0.7		30005202
'Manarq 	0.8		30005201
'Tar 	0.8		30005199
'Pakhshi 	0.8		30005198
'Synchelle 	0.9		30005015
'Kassigainen 	0.9		30002761
'Hatakani 	0.9		30002764
'Sivala 	0.6		30002765
'Uedama 	0.5		30002768
'Haatomo 	0.6		30002800
'Suroken 	0.7		30002801
'Kusomonmon 	0.8		30002802
'Urlen 	1		30000139
'Perimeter 	1		30000144
'Jita 			30000142


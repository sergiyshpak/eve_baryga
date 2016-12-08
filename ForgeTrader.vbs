'URL="https://eve-central.com/home/tradefind_display.html?qtype=Regions&newsearch=1&fromt=10000002&to=10000002"
URL="https://eve-central.com/home/tradefind_display.html?qtype=SystemToRegion&newsearch=1&fromt=30000142&to=10000002"

htmlName="TheForgeTrade.html"

' https://eve-central.com/home/tradefind_display.html?qtype=Regions&newsearch=1&fromt=10000002&to=10000002
' https://eve-central.com/home/tradefind_display.html?qtype=Regions&fromt=10000002&to=10000002&age=24&minprofit=500000&size=8000&startat=50&sort=jprofit
' https://eve-central.com/home/tradefind_display.html?set=1&fromt=10000002&to=10000002&qtype=Regions&age=24&minprofit=500000&size=8000&limit=50&sort=jumps&prefer_sec=1
' https://eve-central.com/home/tradefind_display.html?set=1&fromt=10000002&to=10000002&qtype=Regions&age=24&minprofit=500000&size=8000&limit=100&sort=jumps&prefer_sec=1
'
'https://eve-central.com/home/quicklook.html?typeid=3687
 
On Error Resume Next
 
set xmlhttp = createobject ("msxml2.xmlhttp.3.0")


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile(htmlName,True)  

resFile.write ("<html><head><script src=sorttable.js></script></head><body><table  border=1 class=sortable>" & vbCrLf)
resFile.write ("<tr> <th>from</th> <th>to</th> <th>jumps</th> <th>what</th>  <th>sell</th>" )
resFile.write ("<th>buy</th> <th>count</th><th>polnoe</th> <th>tot money</th> <th>tot prof</th> <th>perc prof</th>  </tr> " & vbCrLf)

' https://eve-central.com/home/tradefind_display.html?set=1&fromt=30000142&to=10000002&qtype=SystemToRegion&age=24&minprofit=100000&size=10000&limit=5000&sort=jprofit&prefer_sec=1

URL1="https://eve-central.com/home/tradefind_display.html?set=1&fromt="
URLFROM="30000142"
URL2="&to="
URL3="&qtype=SystemToRegion&age=24&minprofit=100000&size=10000&limit=5000&sort=jprofit&prefer_sec=1"


Dim urls
urls=Split("10000002 10000016 10000033 10000032 10000067")

for j = 0 to UBound(urls)

	URL=URL1+URLFROM+URL2+urls(j)+URL3

	'MsgBox URL
	xmlhttp.open "get", URL, false
	xmlhttp.send
	MyText= xmlhttp.responseText

	startpos=1

	do while startpos>0

	    fromPosSt=InStr(startpos, MyText,"<b>From:</b>")
		'MsgBox fromPosSt
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
		
		if (profit1Num/totalmoneyNum)>CDbl(0.04) then 
		resFile.write ("<tr> <td>"&fromStr &"</td> <td>"&toStr &"</td> <td>"&jumpsStr &"</td> <td>"&itemStr &"</td>  <td>"&sellprStr &"</td>" & _
        "<td>"&buyprStr&"</td> <td>"&unitsStr&"</td><td>"&unitsPolnStr&"</td><td>"&FormatCurrency(totalmoneyStr)&"</td> <td>"&FormatCurrency(profitStr)&"</td> <td>"&FormatPercent(profit1Num/totalmoneyNum)&"</td>  </tr> " & vbCrLf)
		end if
		
		
		startpos=InStr(profitPosEnd, MyText,"<b>From:</b>")
    loop 
	
next

resFile.write ("</table></body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start " + htmlName


URL="https://zkillboard.com/system/30000142/"

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")

Function CheckDrop(Arg1)
	xmlhttp.open "get", Arg1, false
	xmlhttp.send
	dropText= xmlhttp.responseText
	
	droposSt = InStr(1, dropText,"info_kill_dttm")+16
	droposEnd = InStr(droposSt, dropText,"/td>")-1
    datetimekillStr =  Mid(dropText, droposSt, droposEnd-droposSt)	
	spacePos= InStr(1,datetimekillStr," ")
	gutyear=Mid(datetimekillStr,1,10)
	guttime=Right("000" & Mid(datetimekillStr,spacePos+1),5)
	datetimekillStr = gutyear & " " & guttime 
	
    droposSt = InStr(droposEnd, dropText,"<th>Dropped:</th>")+45
	droposEnd = InStr(droposSt, dropText,"/td>")-1
	dropa = Mid(dropText, droposSt, droposEnd-droposSt)
	
	CheckDrop =   datetimekillStr & "</td><td>" & dropa
End Function


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("z_kill.html",True)  

'''''''   crazy shit...
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")    
dateTime.SetVarDate (now())
'MsgBox  "Local Time:  " & dateTime
'MsgBox  "UTC Time: " & dateTime.GetVarDate (false)
utcTime= dateTime.GetVarDate (false)
'MsgBox  "UTC Time: " & FormatDateTime(utcTime,4)
'''

utcTimeStr =  Year(utcTime) & "-" & _
    Right("0" & Month(utcTime),2)  & "-" & _
    Right("0" & Day(utcTime),2)  & " " & _  
    Right("0" & Hour(utcTime),2) & ":" &_
    Right("0" & Minute(utcTime),2)  
	
resFile.write ("<html><head><script src=sorttable.js></script></head><body>"& utcTimeStr & " -- current time <br><table  border=1 class=sortable>" & vbCrLf)
resFile.write ("<tr> <th>time</th> <th>loot</th> <th>link</th> <th>ship</th>  <th>sys</th> <th>reg</th>  </tr> " & vbCrLf)

Dim urls(3)
urls(0) ="https://zkillboard.com/system/30000142/"
urls(1) ="https://zkillboard.com/system/30000144/"
urls(2) ="https://zkillboard.com/system/30000156/"
urls(3) ="https://zkillboard.com/system/30001363/"     'sobaseki

for j = 0 to UBound(urls)
	URL=urls(j)

	xmlhttp.open "get", URL, false
	xmlhttp.send
	MyText= xmlhttp.responseText

	startpos=1
	for i =1 to 50
		'  winwin killListRow
		curpos=InStr(startpos, MyText,"winwin killListRow")
		
		cur1st=InStr(curpos, MyText,"window.location=")
		kilaStr= Mid(MyText, cur1st+17,15)
		kilaNum= Mid(MyText, cur1st+23,8)
		kilaLink= "https://zkillboard.com"+Mid(MyText, cur1st+17,15)
		timekill =Mid(MyText, cur1st+71,5)

		cur2st=InStr(cur1st,MyText, "<a href="+chr(34)+kilaStr+chr(34))
		cur2end=InStr(cur2st+6, MyText, "<")
		utrata=Mid(MyText, cur2st+26, cur2end-cur2st-26)
		
		cur2ast=InStr(cur2st,MyText, "class="+chr(34)+"eveimage img-rounded"+chr(34)+" alt="  )+1
		cur2aend=InStr(cur2ast+6, MyText, "/")
		karabl=Mid(MyText, cur2ast+33, cur2aend-cur2ast-34)
		
		cur3st=InStr(cur2end,MyText, "/system/")
		cur3end=InStr(cur3st+8,MyText, "<")
		system=Mid(MyText, cur3st+19, cur3end-cur3st-19)

		cur4st=InStr(cur3end,MyText, "/region/")
		cur4end=InStr(cur4st+8,MyText, "<")
		region=Mid(MyText, cur4st+19, cur4end-cur4st-19)
		
		drop="pofig"
		if (karabl<>"Capsule") then
		   drop=CheckDrop(kilaLink)
		   'resFile.write (timekill & "," & kilaLink & "," & karabl & "," & utrata & "," & system & "," &  region & "," &  drop & vbCrLf)
		   resFile.write ("<tr><td>" & drop & "</td><td><a href=" & kilaLink & ">" & kilaLink & "</a></td><td>" & karabl & "</td><td>" )
		   resFile.write ( system & "</td><td>" & region & "</td></tr>" )
		   resFile.write (vbCrLf)
		   'Msgbox "Kill:"+kilaNum+"   Karabl:"+karabl+"  Uron:"+utrata+"   system:"+system+"  region:"+region
		end if
		
		startpos=cur4end+1
		
	next
next
resFile.write ("</table></body></html>")
resFile.Close


set shell = WScript.CreateObject("WScript.Shell")
shell.Run "cmd /c  start z_kill.html"

' https://zkillboard.com/region/10000016/   lonetrek
' https://zkillboard.com/region/10000002/   the forge

' https://zkillboard.com/system/30000142/  jita
'  https://zkillboard.com/system/30000144/  perimeter
' https://zkillboard.com/system/30000156/   josameto
' https://zkillboard.com/system/30000182/  Inaya

' https://zkillboard.com/system/30001389/  isanamo


'  https://zkillboard.com/corporation/1000125/   by concord
' 
URL="https://zkillboard.com/system/30000142/"

Function CheckDrop(Arg1)
     CheckDrop = "need to check " & Arg1
End Function


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("z_killborda.csv",True)  

'''''''   crazy shit...
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")    
dateTime.SetVarDate (now())
'MsgBox  "Local Time:  " & dateTime
'MsgBox  "UTC Time: " & dateTime.GetVarDate (false)
utcTime= dateTime.GetVarDate (false)
'MsgBox  "UTC Time: " & FormatDateTime(utcTime,4)
'''

resFile.write (FormatDateTime(utcTime,4) & ",link,ship,bablo,system,region,drop" & vbCrLf)

set xmlhttp = createobject ("msxml2.xmlhttp.3.0")
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
	'MsgBox kilaLink
	'MsgBox kilaNum
	'MsgBox kilaStr
	
	timekill =Mid(MyText, cur1st+71,5)

	cur2st=InStr(cur1st,MyText, "<a href="+chr(34)+kilaStr+chr(34))
	cur2end=InStr(cur2st+6, MyText, "<")
	utrata=Mid(MyText, cur2st+26, cur2end-cur2st-26)
	
	cur2ast=InStr(cur2st,MyText, "class="+chr(34)+"eveimage img-rounded"+chr(34)+" alt="  )+1
	cur2aend=InStr(cur2ast+6, MyText, "/")
	karabl=Mid(MyText, cur2ast+33, cur2aend-cur2ast-34)
	'MsgBox  karabl
	
	cur3st=InStr(cur2end,MyText, "/system/")
	cur3end=InStr(cur3st+8,MyText, "<")
	system=Mid(MyText, cur3st+19, cur3end-cur3st-19)
	

	cur4st=InStr(cur3end,MyText, "/region/")
	cur4end=InStr(cur4st+8,MyText, "<")
	region=Mid(MyText, cur4st+19, cur4end-cur4st-19)
	
	drop="pofig"
	if (karabl<>"Capsule") then
	   drop=CheckDrop(kilaLink)
	end if

	resFile.write (timekill & "," & kilaLink & "," & karabl & "," & utrata & "," & system & "," &  region & "," &  drop & vbCrLf)
	
	'Msgbox "Kill:"+kilaNum+"   Karabl:"+karabl+"  Uron:"+utrata+"   system:"+system+"  region:"+region
	
	startpos=curpos+1
	
next


resFile.Close
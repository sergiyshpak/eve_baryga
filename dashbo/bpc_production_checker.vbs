 Dim IE
 Set IE = CreateObject("InternetExplorer.Application")
 IE.Visible = 1 

set objFSO=CreateObject("Scripting.FileSystemObject")
Set resFile = objFSO.CreateTextFile("resik.csv",True)
resFile.write("figovina,pribbilo,unitPrice,meterialo,installCost,inventionCost"&vbCrLf  )

itemsStr="2488 1013 2176 2186 2196 2206"
Dim items
items = Split (itemsStr)

for j = 0 to UBound(items)
 item=items(j)
 URLO="https://www.fuzzwork.co.uk/blueprint/?typeid="&item
 IE.navigate URLO
 Do While (IE.Busy)
   WScript.Sleep 4000
 Loop
 figovina=IE.Document.getElementByID("nameDiv").textContent
 
 'msgbox IE.Document.getElementByID("unitPrice").textContent
 'msgbox Replace(IE.Document.getElementByID("unitPrice").textContent,",","")
 
 
 unitPrice=CDbl(Replace(IE.Document.getElementByID("unitPrice").textContent,",",""))
 jobCost=CDbl(Replace(IE.Document.getElementByID("jobCost").textContent,",",""))
 installCost=CDbl(Replace(IE.Document.getElementByID("installCost").textContent,",",""))
 inventionCost=CDbl(Replace(IE.Document.getElementByID("inventionCost").textContent,",",""))

 ' msgbox figovina+" "+unitPrice+" "+jobCost+" "+installCost+" "+inventionCost 
 
 pribu=unitPrice*10- (jobCost+installCost)*10-inventionCost
 resFile.write(figovina&","&CStr(pribu)&","&CStr(unitPrice)&","&CStr(jobCost)&","&CStr(installCost)&","&CStr(inventionCost)&","&URLO&vbCrLf  )
 
next 
 
 
 resFile.close
 
 ie.quit

'http://api.eve-central.com/api/quicklook?typeid=34

'Dim o
'Set o = CreateObject("MSXML2.XMLHTTP")
'o.open "GET", "http://api.eve-central.com/api/quicklook?typeid=34", False
'o.send


set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
'xmlDoc.load("market.xml")
xmlDoc.load("http://api.eve-central.com/api/quicklook?typeid=34")

Set objFSO=CreateObject("Scripting.FileSystemObject")

Set resFile = objFSO.CreateTextFile("market_sell.csv",True)  
resFile.Write "order,region,security,station,station_name,price,vol_remain.text,min_volume.text" & vbCrLf

strQuery="/evec_api/quicklook/sell_orders/order"
Set colNodes = xmlDoc.selectNodes( strQuery )
For Each objNode in colNodes
    resFile.write ("sell,")
	Set region=objNode.getElementsByTagName("region")(0)
	Set station=objNode.getElementsByTagName("station")(0)
	Set station_name=objNode.getElementsByTagName("station_name")(0)
	Set price=objNode.getElementsByTagName("price")(0)
	Set vol_remain=objNode.getElementsByTagName("vol_remain")(0)
	Set min_volume=objNode.getElementsByTagName("min_volume")(0)
	Set security=objNode.getElementsByTagName("security")(0)
	
	resFile.write ( region.text & "," & security.text & "," & station.text & "," & station_name.text & "," & price.text & "," & vol_remain.text & "," & min_volume.text  & vbCrLf)
Next
resFile.Close

Set resFile = objFSO.CreateTextFile("market_buy.csv",True)  
resFile.Write "order,region,security,station,station_name,price,vol_remain,min_volume" & vbCrLf

strQuery="/evec_api/quicklook/buy_orders/order"
Set colNodes = xmlDoc.selectNodes( strQuery )
For Each objNode in colNodes
    resFile.write ("buy,")
	Set region=objNode.getElementsByTagName("region")(0)
	Set station=objNode.getElementsByTagName("station")(0)
	Set station_name=objNode.getElementsByTagName("station_name")(0)
	Set price=objNode.getElementsByTagName("price")(0)
	Set vol_remain=objNode.getElementsByTagName("vol_remain")(0)
	Set min_volume=objNode.getElementsByTagName("min_volume")(0)
	Set security=objNode.getElementsByTagName("security")(0)
	
	resFile.write ( region.text & "," & security.text & "," & station.text & "," & station_name.text & "," & price.text & "," & vol_remain.text & "," & min_volume.text  & vbCrLf)
Next
resFile.Close



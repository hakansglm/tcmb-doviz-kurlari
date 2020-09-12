<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
Session.CodePage = 65001
Session.LCID = 1055
Response.Charset = "UTF-8"

Function TCMB_DovizKuruCek(strDovizTipi)
	Set kurlar = Server.CreateObject("msxml2.DOMDocument" )
		kurlar.async = false
		kurlar.resolveExternals = false
		kurlar.setProperty "ServerHTTPRequest" ,true
		kurlar.load("http://www.tcmb.gov.tr/kurlar/today.xml" )
		Set sonuc =kurlar.getElementsByTagName("Currency" )
			USDA=sonuc.item(0).childnodes.item(3).nodeTypedValue
			USDS=sonuc.item(0).childnodes.item(4).nodeTypedValue
			EURA=sonuc.item(3).childnodes.item(3).nodeTypedValue
			EURS=sonuc.item(3).childnodes.item(4).nodeTypedValue 
		Set sonuc = nothing
	Set kurlar = nothing
	
	Select Case(strDovizTipi)
		Case "USD"
		TCMB_DovizKuruCek = Replace(USDS, "." , "," , 1, -1, 1)
		Case "EUR"
		TCMB_DovizKuruCek = Replace(EURS, "." , "," , 1, -1, 1)
		Case Else
		TCMB_DovizKuruCek = "N/A"
	End Select
	
End Function

'Test
response.write "Dolar Satış: "&TCMB_DovizKuruCek("USD")
response.write "<br>"
response.write "Euro Satış: "&TCMB_DovizKuruCek("EUR")
 %>

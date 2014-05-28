<META HTTP-EQUIV="Content-Type" CONTENT="text/html"; charset="utf-8">
<html>
<head>
<%

'http://10.66.1.5/intranetrk/proyectos/reporte_erp_unilever_pedimentos.asp
'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/unilever-reporte-operaciones-032011.asp

Response.Buffer = TRUE
response.Charset = "utf-8"
Response.Addheader "Content-Disposition", "attachment; filename=ReporteDinamico.xls"
Response.ContentType = "application/vnd.ms-excel"

dim oficina,cvesoficina,validacion
oficina="RKU"
cvesoficina=""
validacion=""

'oficina=Request.QueryString("ofi")
tipope="I" 'Request.QueryString("tipope")
'det=Request.QueryString("det")
'mes=Request.QueryString("mes")
finicio="2011-10-01" 'Request.QueryString("finicio")
ffinal="2011-10-31" 'Request.QueryString("ffinal")

dim strHTML 
strHTML = ""

Server.ScriptTimeOut=100000
%>
<title> Reporte1.. </title>
</head>
<body>
<table width="984" border="2" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666"> 
<tr  width="100" height="100"><td><div><img src="file:///C|/unilever.JPG" width="100" height="98"></img></div></td></tr>
<tr align="center" border="2" bordercolor="#999999" bgcolor="#FFFFFF">
<% genera_registros det,tipope %>
</table>
</body>
</html>
<%

sub genera_registros(det,tipope)
dim c
 c=chr(34)

 %><td width="100" align="center" bgcolor="#79BCFF"><div align="center"><font color="#666666"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Número de Pedimento </font></font></div></td>
<%
'genera_html "e","DIVISION ","center"   '
genera_html "e","Patente Aduanal","center"   '2
genera_html "e","Clave de la Aduana	","center"   '3
genera_html "e","Fecha de Pago","center"   '4
genera_html "e","Valor de la Mercancía en Aduana","center"   '5
genera_html "e","Importe del IVA de la Mercancía en Aduana","center"   '6
genera_html "e","Pais de procedencia","center"   '7
genera_html "e","Pais de Origen","center"   '8
genera_html "e","Pais Destino","center"   '9
genera_html "e","Nombre del Agente Aduanal","center"   '10
genera_html "e","NOMBRE DEL PROVEEDOR o CLIENTE","center"   '11
genera_html "e","ID  FISCAL DEL PROVEEDOR o CLIENTE","center"   '12
genera_html "e","ID  FISCAL DEL PROVEEDOR o CLIENTE","center"   '13
genera_html "e","CLASE  DE  DE PEDIMENTO","center"   '14




%></tr><%

sqlAct= "select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'IMPO' as '14' ," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from rku_extranet.ssdagi01 as i  " & _
"   inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join rku_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'EXPO' as '14' ," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from rku_extranet.ssdage01 as i  " & _
"   inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join rku_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'IMPO' as '14' ," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from sap_extranet.ssdagi01 as i  " & _
"   inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join sap_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'EXPO' as '14' ," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from sap_extranet.ssdage01 as i  " & _
"   inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join sap_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'IMPO' as '14'," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from tol_extranet.ssdagi01 as i  " & _
"   inner join tol_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join tol_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join tol_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join tol_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'EXPO' as '14' ," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from tol_extranet.ssdage01 as i  " & _
"   inner join tol_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join tol_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join tol_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join tol_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'IMPO' as '14'," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from dai_extranet.ssdagi01 as i  " & _
"   inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join dai_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'EXPO' as '14'," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from dai_extranet.ssdage01 as i  " & _
"   inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join dai_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'IMPO' as '14'," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from lzr_extranet.ssdagi01 as i  " & _
"   inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join lzr_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' " & _
" union all " & _
"select " & _
" i.numped01 as '1', " & _
" i.patent01 as '2', " & _
" i.adusec01 as '3', " & _
" i.fecpag01 as '4', " & _
" i.refcia01 as '5', " & _
" cf3.import36 as '6', " & _
"  r.ptoemb01 as '177',  " & _
"  r.cveptoemb as '17',  " & _
" '' as '7', " & _
" i.cvepod01 as '8', " & _
" '' as '9', " & _
" i.patent01 as '10', " & _
" prv.nompro22 as '11', " & _
" prv.irspro22 as '12', " & _
" prv.irspro22 as '13', " & _
" 'EXPO' as '14'," & _
" i.cveped01 as '15'," & _
" i.refcia01 as '29' " & _
"   from lzr_extranet.ssdage01 as i  " & _
"   inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join lzr_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
"            left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"                    left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and i.fecpag01 >= '"& finicio &"' and i.fecpag01 <='"& ffinal &"' "





'response.Write(sqlAct)
'response.end()


Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()

dim ref,refAux
refAux=""

while not act2.eof
response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")


'genera_html "e","Patente Aduanal","center"   '2
'genera_html "e","Clave de la Aduana	","center"   '3
'genera_html "e","Fecha de Pago","center"   '4
'genera_html "e","Valor de la Mercancía en Aduana","center"   '5
'genera_html "e","Importe del IVA de la Mercancía en Aduana","center"   '6
'genera_html "e","Pais de procedencia","center"   '7
'genera_html "e","Pais de Origen","center"   '8
'genera_html "e","Pais Destino","center"   '9
'genera_html "e","Nombre del Agente Aduanal","center"   '10
'genera_html "e","NOMBRE DEL PROVEEDOR o CLIENTE","center"   '11
'genera_html "e","ID  FISCAL DEL PROVEEDOR o CLIENTE","center"   '12
'genera_html "e","CLASE  DE  DE PEDIMENTO","center"   '13
'genera_html "e","País de Origen ","center"   '14


genera_html "d",act2.fields("1").value,"center"  'Numero Pedimento
genera_html "d",act2.fields("2").value,"center"  'Patente
genera_html "d",act2.fields("3").value,"center"  'Cve aduana
genera_html "d",act2.fields("4").value,"center"  'Fecha Pago
genera_html "d",retornaValorAduana(act2.fields("5").value,mid(act2.fields("5").value,1,3)),"center"  'Valor Mcia
genera_html "d",act2.fields("6").value,"center"  'Importe IVA
genera_html "d",retornaCampoPuertoEmb(act2.fields("177").value,act2.fields("17").value,"cvepai01",mid(act2.fields("29").value,1,3)),"center"  'País de Procedencia
genera_html "d",retornaPaisOrigen(act2.fields("5").value,mid(act2.fields("5").value,1,3)),"center"  'Pais de Origen
genera_html "d",retornaPaisDestino(act2.fields("5").value,mid(act2.fields("5").value,1,3)),"center"  'Pais Destino
genera_html "d",retornaAgenteAduanal(act2.fields("10").value),"center"  'Agente Aduanal
genera_html "d",act2.fields("11").value,"center"  'Nombre Cte o Prov.
genera_html "d",act2.fields("12").value,"center"  'ID Fiscal del cte o proveedor
genera_html "d",act2.fields("13").value,"center"  'ID Fiscal del cte o proveedor
genera_html "d",act2.fields("14").value&" CVE:"&act2.fields("15").value,"center"  'Tipo de pedimento


'genera_html "d",retornaCampoPuertoEmb(act2.fields("177").value,act2.fields("17").value,"cvepai01",mid(act2.fields("29").value,1,3)),"center"  'País de Procedencia
'genera_html "d",retornaRegion(act2.fields("14").value),"center"  'Region

'genera_html "d",retornaIMPORTADOR(act2.fields("21").value,mid(act2.fields("29").value,1,3)),"center"  'IMPORTADOR
'genera_html "d",act2.fields("22").value,"center"  'Cantidad 
'genera_html "d",retornaAduana(act2.fields("26").value),"center"  'Aduana
'genera_html "d",retornaAgenteAduanal(act2.fields("27").value),"center"  'Agente Aduanal
'genera_html "d",act2.fields("28").value,"center"  'Patente Agente Aduanal
'genera_html "d",act2.fields("29").value,"center"  'No. De Trafico
'genera_html "d",retornaCampoContenedores(act2.fields("29").value,"marc01",mid(act2.fields("29").value,1,3)),"center"  'No de Contenedor 
'genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
'genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
'genera_html "d",act2.fields("33").value,"center"  'Mes
'genera_html "d",act2.fields("34").value,"center"  'No.Semana
'genera_html "d",act2.fields("35").value,"center"  'Cantidad de Operaciones 
'genera_html "d",retornaCantContenedores(act2.fields("29").value,"'ISO','CON'",mid(act2.fields("29").value,1,3)),"center"  'Cantidad de Contenedores
'genera_html "d",retornaCantContenedores(act2.fields("29").value,"'BUL','CAJ','BID','PAL'",mid(act2.fields("29").value,1,3)),"center"  'PALLETS/BULTOS 
'genera_html "d",retornaTipoContenedores(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  'TIPO DE CONTENEDOR/ CAJA
'genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
'genera_html "d",act2.fields("40").value,"center"  'Fecha BL
'genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
'genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento

'genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
'genera_html "d",retornaMontoAnticipo(act2.fields("29").value,"ANT",mid(act2.fields("29").value,1,3)),"center"  ' Monto de Anticipo 

   
 '  genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
 '  genera_html "d",act2.fields("77").value,"center"  'IVA %
 '  genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
 '  genera_html "d",act2.fields("79").value,"center"  ' PREVAL. 
 '  genera_html "d",sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  ' TOTAL IMPUESTOS 
 '  genera_html "d",(sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))/act2.fields("65").value),"center"  ' Total Impuestos USD  
   
'	genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA) 
'	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' DEMORAS 
'	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' ESTADIAS 
'	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' MANIOBRAS  
'	genera_html "d","?","center"  ' ALMACENAJES 
'	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"OTROS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' OTROS 
	'genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)),"center"  ' TOTAL GASTOS DIVERSOS 
	
'	genera_html "d",(retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)) /act2.fields("65").value),"center"  ' TOTAL GASTOS DIVERSOS USD 
'	genera_html "d",retornaHonorarios(act2.fields("29").value,"chon31",mid(act2.fields("29").value,1,3)),"center"  ' HONORARIOS AG AD. $ 





 response.Write("</tr>")

 act2.movenext()
wend
response.End()

end sub

sub genera_html(tipo,valor,alineacion)
 if(tipo = "e")then
  response.Write("<td width="&c&"94"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#79BCFF"&c&"><div align="&c&alineacion&c&"><font size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
 else 
  response.Write("<td width="&c&"94"&c&" align="&c&alineacion&c&" nowrap><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
 end if

end sub

function regresa_fecha_cuenta_gastos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select r.refe31,min(cta.fech31) as fech31 from e31cgast as cta, d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"'  "&_
" group by r.refe31"

Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
if not(act2.eof) then
 regresa_fecha_cuenta_gastos =act2.fields("fech31").value
else
  regresa_fecha_cuenta_gastos =valor
   end if
end function

function codigoProveedor(desc)
dim res,desc2
res = "no"
Path=Server.MapPath("catprobd.xls")
desc2=replace(desc," ","%")
Set ConexionBD = Server.CreateObject("ADODB.Connection") 
ConexionBD.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & Path
Set rsVac = Server.CreateObject("ADODB.Recordset") 
rsVac.Open "Select * From A1:B50 where descpro like '" & desc2 & "'", ConexionBD,3,3 

if not(rsVac.eof)then
 res =rsVac.fields("cvepro") '&","&rsVac.fields("descpro")
else
 res = desc
end if 

codigoProveedor = res
end function

function retornaCampoCtaGastos(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = "select r."& campo &" as campo from "&oficina&"_extranet.e31cgast as cta " &_
 " inner join  "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
 " where  r.refe31 = '"& referencia &"' and cta.esta31 <> 'C' "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoCtaGastos = valor
 else
  retornaCampoCtaGastos =valor
 end if
end function


function retornaMontoAnticipo(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = " select sum(dm.mont11) as campo " & _
			" from "&oficina&"_extranet.d11movim as dm " & _
			" where dm.refe11='"& referencia&"' and dm.conc11 = '"&campo&"' "
	'and dm.cgas11 <> ''"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaMontoAnticipo = valor
 else
  retornaMontoAnticipo =valor
 end if
end function


function retornaIMPORTADOR(clave,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = "select c.nomcli18 as campo from "&oficina&"_extranet.ssclie18 as c where c.cvecli18 = "&clave

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaIMPORTADOR = valor
 else
  retornaIMPORTADOR =valor
 end if
end function







function retornaCampoPuertoEmb(pto,val,campo,oficina)

dim c,valor
 c=chr(34)
 valor=""
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
 
 if (ucase(oficina) = "SAP")then
	val=0
	pto=""
 end if
 
 if(pto <> "") then 

	sqlAct="SELECT "& campo &" as campo FROM "&oficina&"_extranet.c01ptoemb where cvepto01 ="& val &" and nompto01 like concat(replace('"& pto &"','''',''),'%')"
	'response.Write(sqlAct)
	'response.End()

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	 if not(act2.eof) then
	 valor = act2.fields("campo").value
	 act2.movenext()
	 while not act2.eof
	   valor = valor&", "&act2.fields("campo").value
	   act2.movenext()
	 wend
	  retornaCampoPuertoEmb = valor
	 else
	  retornaCampoPuertoEmb =valor
	end if
 else
	retornaCampoPuertoEmb =valor	  
 end if
end function

function retornaAgenteAduanal(valor)
dim val
val = ""
if(valor = "3921")then
 val = "Luis E. de la Cruz Reyes"
else
if(valor = "3210")then
 val = "Rolando Reyes Kuri"
else
if(valor = "3945")then
 val = "Jesús Gómez Reyes"
else
if(valor = "3931")then
 val = "Sergio Alvarez Ramírez"
else
if(valor = "3044")then
 val = "Carlos Humberto Zesati Andrade"
else
val =""
end if
end if
end if
end if
end if

retornaAgenteAduanal = val
end function


function retornaAduana(valor)
dim val
val = ""
if(valor = "470")then
 val = "México"
else
if(valor = "430")then
 val = "Veracruz"
else
if(valor = "810")then
 val = "Altamira"
else
if(valor = "160")then
 val = "Manzanillo"
else
if(valor = "510")then
 val = "Lázaro Cardenas"
else
if(valor = "650")then
 val = "Toluca"
else
val =""
end if
end if
end if
end if
end if
end if

retornaAduana = val
end function




function codigoCliente(desc)
dim res,desc2
res = "no"
Path=Server.MapPath("catclibd.xls")
desc2=replace(desc," ","%")
Set ConexionBD = Server.CreateObject("ADODB.Connection") 
ConexionBD.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & Path
Set rsVac = Server.CreateObject("ADODB.Recordset") 
rsVac.Open "Select * From A1:B35 where desccli like '" & desc2 & "'", ConexionBD,3,3 
 'response.write("Select * From A1:B15 where desccli like '%" & desc2 & "%'")
 'response.End()
if not(rsVac.eof)then
 res =rsVac.fields("cvecli")  '&","&rsVac.fields("desccli")

else
 res = desc
end if 

codigoCliente = res
end function


function retornaHonorarios(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

sqlAct=" select cta."&campo&" as campo from "&oficina&"_extranet.e31cgast as cta  " & _
       " inner join "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
       " where  r.refe31 = '"& referencia &"' and cta.esta31 = 'I' "



Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaHonorarios = valor
 else
  retornaHonorarios =valor
 end if
end function

function retornaConceptosPH(oficina,topico)
dim cad
cad = "NA"

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if oficina = "SAP" then

if topico = "ALMACENAJES-MANIOBRAS" then
  cad= "2,3,63,65,86,111,111,112,142,174,181,183,183,186,188,189,190,196,196,208,209,210,211,212,214,216,218,234,251,256,258,265,269,284,286,287,288,289,290,291,292,293,294,295,296,297,298,299,300,301,303,304,305,307,309,323,331,336,351"
end if
if topico = "DEMORAS" then
  cad= "6,14,46,63,129,156,352"
end if
if topico = "ESTADIAS" then
  cad="144"
end if
if topico = "OTROS" then
  cad="306,313,350,351,352"
end if


else 
  if oficina = "CEG" then
  
    if topico = "ALMACENAJES-MANIOBRAS" then
     cad= "2,4,59,77,100,223,223,235,235,241"
    end if
	if topico = "DEMORAS" then
	 cad= "11,48,99,150"
	end if
	if topico = "ESTADIAS" then
	 cad="NA"
	end if
	if topico = "OTROS" then
	 cad="239"
    end if
  
  else 
     if oficina = "TOL" then
	 
	    if topico = "ALMACENAJES-MANIOBRAS" then
		 cad= "2,2,10,127,128"
		end if
		if topico = "DEMORAS" then
		 cad= "79"
		end if
		if topico = "ESTADIAS" then
		 cad="123"
		end if
		if topico = "OTROS" then
		 cad="NA"
		end if
		
     else 
	   if oficina = "LZR" then
	         if topico = "ALMACENAJES-MANIOBRAS" then
			 cad= "4,78,115,116,119,125,160,167,167,203,230,230,244,297,312"
			end if
			if topico = "DEMORAS" then
			 cad= "11"
			end if
			if topico = "ESTADIAS" then
			 cad="77"
			end if
			if topico = "OTROS" then
			 cad="NA"
			end if
       else 
	       if oficina = "RKU" then
		          if topico = "ALMACENAJES-MANIOBRAS" then
					 cad= "2,4,78,115,116,119,125,160,167,167,203,230,230,244,297,304,311,312,313,359,359"
					end if
					if topico = "DEMORAS" then
					 cad= "11,310,376"
					end if
					if topico = "ESTADIAS" then
					 cad="77"
					end if
					if topico = "OTROS" then
					 cad="341"
					end if
           else 
		       if oficina = "DAI" then
			           if topico = "ALMACENAJES-MANIOBRAS" then
						 cad= "2,2,10,93,127,128,155,163,163,166,170"
						end if
						if topico = "DEMORAS" then
						 cad= "79,171"
						end if
						if topico = "ESTADIAS" then
						 cad="123"
						end if
						if topico = "OTROS" then
						 cad="NA"
						end if
               else 
   			          cad = "NE"
               end if
           end if
       end if
     end if
  end if
end if
retornaConceptosPH = cad
end function

function retornaPagosHechos(referencia,conceptos,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=""

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

'sqlAct=" select i.refcia01 as Ref,(ep.mont21*if(ep.deha21 = 'C',-1,1)) as Importe  " & _
sqlAct =" select i.refcia01 as Ref,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe " & _
		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and ep.fech21 = dp.fech21 and ep.conc21 in ("&conceptos&") and ep.esta21 <> 'S' and ep.esta21 <> 'C' " & _
		" where i.rfccli01 = 'UME651115N48'   and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''   group by i.refcia01 "

'response.Write(sqlAct)
'response.End()

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("Importe").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("Importe").value
   act2.movenext()
 wend
  retornaPagosHechos = valor
 else
  retornaPagosHechos =valor
 end if
 else
   retornaPagosHechos ="ERROR"
 end if

end function


function retornaTOTALPagosHechos(referencia,tipope,oficina)
dim c,valor
 c=chr(34)
 valor="0"

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

sqlAct =" select i.refcia01 as Ref,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe " & _
		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and ep.fech21 = dp.fech21 and ep.esta21 <> 'S' and ep.esta21 <> 'C' " & _
		" where i.rfccli01 = 'UME651115N48'   and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''  group by i.refcia01 "

'response.Write(sqlAct)
'response.End()

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("Importe").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("Importe").value
   act2.movenext()
 wend
  retornaTOTALPagosHechos = valor
 else
  retornaTOTALPagosHechos = valor
 end if
 else
   retornaTOTALPagosHechos ="ERROR"
 end if

end function

function retornaCampoContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select "& campo &" as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"'  "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoContenedores = valor
 else
  retornaCampoContenedores =valor
 end if
end function


function retornaCantContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
sqlAct="select count(*) as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"' and clas01 in ("&campo&") "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCantContenedores = valor
 else
  retornaCantContenedores =valor
 end if
end function


function DatosContenedor(Tipo)
dim val
val = Tipo
if(Tipo = "1")then
  val="Contenedor Estandar 40 pulg (Standard Container 40 pulg)"
end if

if(Tipo = "2")then
  val="Contenedor Estandar 40 pulg (Standard Container 40 pulg)"
end if

if(Tipo = "3")then
  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
end if

if(Tipo = "17")then
val="Contenedor Refrigerante Cubo Alto 17 pulg (High Cube Refrigerated Container 40 pulg"
end if
DatosContenedor = val
end function

function retornaTipoContenedores(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select distinct cn4.tipcon40 as campo from "& oficina &"_extranet.sscont40 as cn4 where cn4.refcia40 = '"& referencia &"' "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = DatosContenedor(act2.fields("campo").value)
 act2.movenext()
 while not act2.eof
   valor = DatosContenedor(valor) &", "& DatosContenedor(act2.fields("campo").value)
   act2.movenext()
 wend
  retornaTipoContenedores = valor
 else
  retornaTipoContenedores =valor
 end if
end function


function retornaDivision(clave)
dim val
val= ""
if clave = "11000" then
val = "Centro de Distribución"
end if
if clave = "11001" then
val = "Planta Helados"
end if
if clave = "11002" then
val = "Planta Foods"
end if
if clave = "11003" then
val = "Planta HPC"
end if
if clave = "11004" then
val = "Todas"
end if

retornaDivision = val
end function


function retornaPlantaEntrega(clave)
dim val
val= ""
if clave = "11000" then
val = "CDU"
end if
if clave = "11001" then
val = "TULTITLAN"
end if
if clave = "11002" then
val = "LERMA"
end if
if clave = "11003" then
val = "CIVAC"
end if
if clave = "11004" then
val = "ESPECIALES"
end if

retornaPlantaEntrega = val
end function


function sumaTotalImpuestos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="0"
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct=" select sum(import36) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('1','3','6','15')   and refcia36 = '"&referencia&"' "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  sumaTotalImpuestos = valor
 else
  sumaTotalImpuestos =valor
 end if

'ADV/IGI+DTA+IVA+PREVAL

end function

function retornaRegion(clave)
dim val
val= ""
if clave = "ARG" then
val = "SUDAMERICA"
end if
if clave = "BRA" then
val = "SUDAMERICA"
end if
if clave = "CHL" then
val = "SUDAMERICA"
end if
if clave = "COL" then
val = "SUDAMERICA"
end if
if clave = "DEU" then
val = "EUROPA"
end if


if clave = "DOM" then
val = "ANTILLAS"
end if
if clave = "ESP" then
val = "EUROPA"
end if
if clave = "FRA" then
val = "EUROPA"
end if
if clave = "GBR" then
val = "EUROPA"
end if
if clave = "SLV" then
val = "CENTROAMERICA"
end if

if clave = "THA" then
val = "ASIA"
end if
if clave = "USA" then
val = "NORTE AMERICA"
end if
if clave = "ZYA" then
val = "EUROPA"
end if

if clave = "CAN" then
val = "NORTE AMERICA"
end if

if clave = "CHN" then
val = "ASIA"
end if

retornaRegion = val
end function


function retornaValorAduana(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="0"
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct=" select sum(vaduan02) as campo from "&oficina&"_extranet.ssfrac02 where refcia02 = '"&referencia&"'"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  retornaValorAduana = valor
 else
  retornaValorAduana =valor
 end if


end function

function retornaPaisOrigen(referencia,oficina)
dim valor

if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
sqlAct=" select refcia01 ,cvepod01 as campo from "&oficina&"_extranet.ssdagi01 where refcia01 = '"&referencia&"'"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  retornaPaisOrigen = valor
 else
  retornaPaisOrigen =valor
 end if


end function


function retornaPaisDestino(referencia,oficina)
dim valor

if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
sqlAct=" select refcia01 ,cvepod01 as campo from "&oficina&"_extranet.ssdage01 where refcia01 = '"&referencia&"'"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  retornaPaisDestino = valor
 else
  retornaPaisDestino =valor
 end if


end function

%>


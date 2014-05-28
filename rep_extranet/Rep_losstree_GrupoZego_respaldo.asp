<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000


strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
'permi = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")
if not permi = "" then
	permi = "  and (" & permi & ") "
end if
AplicaFiltro = False
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")


Tiporepo = Request.Form("TipRep")

if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	blnAplicaFiltro = true
end if
if blnAplicaFiltro then
	permi = " AND cvecli01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	permi = ""
end if

if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	oficina_adu=GAduana
	jnxadu=Session("GAduana")

	select case jnxadu
		case "VER"
			strOficina="rku"
		case "MEX"
			strOficina="dai"
		case "MAN"
			strOficina="sap"
		case "GUA"
			strOficina="rku"
		case "TAM"
			strOficina="ceg"
		case "LAR"
			strOficina="LAR"
		case "LZR"
			strOficina="lzr"
	end select

	cve=request.form("cve")
	ofi=request.form("ofi")
	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	Vrfc=Request.Form("rfcCliente")
	Vckcve=Request.Form("ckcve")
	Vclave=Request.Form("txtCliente")
	'response.write(Vrfc & " | ")
	'response.write(Vckcve & " | ")
	'response.write(Vclave & " | ")
	'response.end()
	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & Diai

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	nocolumns = 16
	tablamov = ""
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write("DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	
	 'Response.Write(GeneraSQL)
	' Response.End()
	
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(GeneraSQL)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left"">" &_
												"REPORTE ÁRBOL DE PERDIDAS.." &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
												"Del " & fi & " Al " & ff &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("Tipo de Operación") &_
								celdahead("Fecha de Despacho") &_
								celdahead("Fecha de Cuenta Gastos") &_
								celdahead("Nombre de Cliente") &_
								celdahead("Clave de Cliente") &_
								celdahead("Observaciones") &_
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Oficina") &_
								celdahead("Proveedor") &_
								celdahead("Descripcion de la Mercancia") &_
								celdahead("Causal") &_
								celdahead("Responsable") &_
								celdahead("Dias Transcurridos (desde el despacho )") & _
								celdahead("Importe Anticipo") &_
								celdahead("Status")
		header = header &	"</tr>"
		
		datos = ""
		Referencia = ""
		ubica = ""
		facturas = ""
		contenedores = ""
		total = ""
		importe = 0
		Do Until RSops.EOF
 
				datos = datos &	"<tr>" &_
								celdadatos(RSops.Fields.Item("Tipo de Operacion").Value) &_
								celdadatos(RSops.Fields.Item("Fecha de Despacho").Value) &_
								celdadatos(RSops.Fields.Item("fech31").Value) &_
								celdadatos(RSops.Fields.Item("nomcli").Value) &_
								celdadatos(RSops.Fields.Item("cvecli").Value) &_
								celdadatos(RSops.Fields.Item("Observaciones").Value) &_
								celdadatos(RSops.Fields.Item("Referencia").Value) &_
								celdadatos(Cstr(RSops.Fields.Item("anopdto").Value) & " " & Cstr(RSops.Fields.Item("cveadu").Value) & " " & Cstr(RSops.Fields.Item("patente").Value) & " " & Cstr(RSops.Fields.Item("pdto").Value)) &_
								celdadatos(RSops.Fields.Item("Oficina").Value) &_
								celdadatos(RSops.Fields.Item("Proveedor").Value) &_
								celdadatos(RSops.Fields.Item("Descripcion de la Mercancia").Value) &_
								celdadatos(RSops.Fields.Item("Causal").Value) &_
								celdadatos(RSops.Fields.Item("Responsable").Value) &_
								celdadatos(RSops.Fields.Item("Dias Transcurridos").Value) & _
								celdadatos(RSops.Fields.Item("mont11").Value) & _
								celdadatos(RSops.Fields.Item("Status").Value)
				datos = datos &	"</tr>"
								
			Rsops.MoveNext()
		Loop
	

	' Response.Write(info & header & datos & "</table><br>" & prom)
	' Response.End()
	html = info & header & datos & "</table><br>"
	
	
	End If
end if


function celdahead(texto)
	cell = "<td bgcolor = ""#006699"" width=""100"" nowrap>" &_
				"<center>" &_
					"<strong>" &_
						"<font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
							texto &_
						"</font>" &_
					"</strong>" &_
				"</center>" &_
			"</td>"
	celdahead = cell
end function

function celdadatos(texto)
'On error resume next
	If IsNull(texto) = True Or texto = "" Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function filtro
	if Vckcve = 0 then
		if Vrfc <> "0" then
			condicion = "AND i.rfccli01 = '" & Vrfc & "' "
		else
			'condicion = "AND i.rfccli01 = 'UME651115N48' "
			condicion = " "
		end if
	else
		if Vclave <> "Todos" Then
			condicion = "AND i.cvecli01 = " & Vclave & " "
		Else
			permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
			condicion = permi
			condicion = "AND " & condicion
			if condicion = "AND cvecli01=0 " Then
				condicion = ""
			end if
		End If
	end if
	filtro = condicion
end function

function GeneraSQL
	SQL = ""
	
	if ofi = "t" then
	'----------------clusular nuevas


	'----------------Clausulas anteriores
						'"AND i.firmae01 IS NOT NULL AND i.firmae01 <> '' " &_
						'"and (i.fecpag01 >=  '" & DateI & "' AND i.fecpag01 <= '" & DateF & "') " &_
						'" AND (bs.Detsit01 = 730 or bs.Detsit01 = 710)  " & _
						'"and (bs.Fechst01 is not null and bs.Fechst01 <> '') " &_
						'"and ((cta.cgas31 IS NULL or cta.cgas31 = '0000000' or cta.cgas31 = '') OR ((cta.fech31 IS NULL OR cta.fech31 = '0000-00-00' OR cta.fech31 = ''))) "  & condicion &_

		
							'"cta.frec31,cta.fech31,cta.cgas31, " & _
						'" d11.fech11,d11.refe11,d11.mont11,d11.conc11,d11.asie11,d11.desc11, " & _
						'" dl11.conc11 as Liquidacion, dl11.mont11 as ImpLiquidacion " & _
	'--------------otro cambio
	'					" IF(cau.c01causa is not null and cau.c01causa <> '', cau.c01causa, " & _
	'					" if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not nu'll ,'2.6 FACTURADA, PENDIENTE DE INGRESAR CON UNILEVER', " & _
	'					" if(bs.fechst01 is null,'1.8 OPERACION PENDIENTE POR DESPACHAR', " & _
	'                    " if(bs.fechst01 is not null and cta.cgas31 is null and (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) <=3,'2.3 EN PROCESO DE FACTURACION MENOR O IGUAL A 3 DIAS','No Capturada'))))  as 'Causal', " & _'
	'					" if(cta.cgas31 is null and bs.fechst01 is null and etx.id_resp is null, 'UNILEVER',( IF(etx.id_resp is not null or etx.id_resp <> '' or etx.id_resp <> 0,   (select group_concat(distinct cres.nom_resp) from rku_status.cat_resp as cres where cres.id_resp = etx.id_resp) , 'GRUPO ZEGO') )) as 'Responsable', " & _

		
		
		condicion = filtro
		
		For i = 0 to 5
		
			Select Case i
					Case 0
						aduanaTmp = "rku"
						adu= "'GRUPO REYES KURI,  S.C.'"
					Case 1
						aduanaTmp = "dai"
						adu = "'DESPACHOS AEREOS INTEGRADOS, S.C.'"
					Case 2
						aduanaTmp = "tol"
						adu = "'COMERCIO EXTERIOR DEL GOLFO, S.C.'"
					Case 3
						aduanaTmp = "sap"
						adu = "'SERVADUANALES DEL PACIFICO, S.C.'"
					Case 4
						aduanaTmp = "lzr"
						adu = "'SERVADUANALES DEL PACIFICO, S.C.'"
					Case 5
						aduanaTmp = "ceg"
						adu = "'COMERCIO EXTERIOR DEL GOLFO, S.C.'"
			End Select
				
			SQL = SQL & "SELECT 'IMPORTACION' as 'Tipo de Operacion', " & chr(13) & chr(10)
					SQL = SQL & "CASE i.cvecli01 " & chr(13) & chr(10)
					SQL = SQL & "WHEN '11000' THEN 'Virginia Leon' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '11001' THEN 'Gilberto Cruz' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '11002' THEN 'Iray Hinojosa' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '11003' THEN 'Francisco Bernal' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '14000' THEN 'Francisco Bernal' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '12000' THEN 'Georgina Perez' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '12001' THEN 'Lucero Bahena' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '12001' THEN 'Roberto Navarrete' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '12002' THEN 'Jorge Islas' " & chr(13) & chr(10)
					SQL = SQL & "WHEN '11004' THEN 'Monserrat Rodriguez' " & chr(13) & chr(10)
				SQL = SQL & "ELSE '' END AS 'Key Account', " & chr(13) & chr(10)
				SQL = SQL & "DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%m-%y') AS 'Fecha de Despacho', " & chr(13) & chr(10)
				SQL = SQL & "c.fdsp01, " & chr(13) & chr(10)
				SQL = SQL & "i.refcia01 AS 'Referencia', " & chr(13) & chr(10)
				SQL = SQL & "i.nomcli01 AS 'nomcli', " & chr(13) & chr(10)
				SQL = SQL & "i.cvecli01 AS 'cvecli', " & chr(13) & chr(10)
				SQL = SQL & "etx.m_observ AS 'Observaciones', " & chr(13) & chr(10)
				SQL = SQL & "cta.cgas31 AS 'Cuenta de Gastos', " & chr(13) & chr(10)
				SQL = SQL & "cta.fech31, " & chr(13) & chr(10)
				SQL = SQL & "cta.fech31 AS 'Fecha CG', " & chr(13) & chr(10)
				SQL = SQL & "cta.frec31 AS 'Fecha de Recepcion', " & chr(13) & chr(10)
				SQL = SQL & "cta.esta31 as 'Estado', " & chr(13) & chr(10)
				SQL = SQL & "DATE_FORMAT(i.fecpag01, '%y') as anopdto, " & chr(13) & chr(10)
				SQL = SQL & "i.cveadu01 as cveadu, " & chr(13) & chr(10)
				SQL = SQL & "i.patent01 as patente, " & chr(13) & chr(10)
				SQL = SQL & "i.numped01 as pdto, " & chr(13) & chr(10)
				SQL = SQL & adu & " as 'Oficina', " & chr(13) & chr(10)
				SQL = SQL & "i.nompro01 as 'Proveedor', " & chr(13) & chr(10)
				SQL = SQL & "fr.d_mer102 AS 'Descripcion de la Mercancia', " & chr(13) & chr(10)
				SQL = SQL & "if(cta.cgas31 is null or cta.cgas31 = '','Despachada pendiente de facturar',if(cta.frec31 is null or cta.frec31 ='' or cta.frec31 = '0000-00-00','Facturada pendiente de entregar','OK')) as 'Status', "  & chr(13) & chr(10)
				SQL = SQL & "IF(cau.c01causa is not null and cau.c01causa <> '', cau.c01causa, if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', if(cta.cgas31 is null and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada'))) as 'Causal', " & chr(13) & chr(10)
				SQL = SQL & "IF(cau.c01causa is not null and cau.c01causa <> '', cre.nom_resp, if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', if(cta.cgas31 is null and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado'))) as 'Responsable', " & chr(13) & chr(10)
				SQL = SQL & "ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', " & chr(13) & chr(10)
				SQL = SQL & "(TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos', " & chr(13) & chr(10)
				SQL = SQL & "IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion', " & chr(13) & chr(10)
				SQL = SQL & "d11.mont11 " & chr(13) & chr(10)
			SQL = SQL & "FROM "&aduanaTmp&"_extranet.ssdagi01 AS i " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.c01refer AS c ON i.refcia01 = c.refe01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.Numpat01 = i.patent01 AND bs.Detsit01 in (730,710) " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_status.c01caus as cau on cau.c01clavec = etx.clavec and (cau.c01tipoc = 'A' and cau.c01tipoo = '0') " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_status.cat_resp as cre on cre.id_resp = etx.id_resp " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 and (cta.frec31 = '0000-00-00' or cta.frec31 is null) and cta.esta31 <> 'C' " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31 and dl11.conc11 = 'LIQ' " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " & chr(13) & chr(10)
			SQL = SQL & "WHERE i.cveped01 <> 'R1' " & chr(13) & chr(10)
				SQL = SQL & "and (cta.fech31 is not null  or d11.conc11 is not null) " & chr(13) & chr(10)
				SQL = SQL & "and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " & chr(13) & chr(10)
				SQL = SQL & "and (bs.Fechst01 <> '0000-00-00' or bs.Fechst01 <> '' or bs.Fechst01 IS NOT NULL) " & chr(13) & chr(10)
				SQL = SQL & "and dl11.conc11 is null " & chr(13) & chr(10)
				SQL = SQL & "and (i.fecpag01 >= '"&DateI&"' and i.fecpag01 <= '"&DateF&"' ) " & chr(13) & chr(10)
				SQL = SQL & condicion & chr(13) & chr(10)
			SQL = SQL & "GROUP BY i.refcia01,cau.c01causa " & chr(13) & chr(10)
			SQL = SQL & "HAVING Estado is null or Estado = 'I' " & chr(13) & chr(10)

			if (i<>5) then
				SQL = SQL & "UNION ALL " & chr(13) & chr(10)
			else
				SQL = SQL & "ORDER BY 19 DESC " & chr(13) & chr(10)
			end if
		Next
		
	else
			Select Case ofi
					Case "r"
						aduanaTmp = "rku"
						adu= "'GRUPO REYES KURI,  S.C.'"
					Case "d"
						aduanaTmp = "dai"
						adu = "'DESPACHOS AEREOS INTEGRADOS, S.C.'"
					Case "t"
						aduanaTmp = "tol"
						adu = "'COMERCIO EXTERIOR DEL GOLFO, S.C.'"
					Case "s"
						aduanaTmp = "sap"
						adu = "'SERVADUANALES DEL PACIFICO, S.C.'"
					Case "l"
						aduanaTmp = "lzr"
						adu = "'SERVADUANALES DEL PACIFICO, S.C.'"
					Case "c"
						aduanaTmp = "ceg"
						adu = "'COMERCIO EXTERIOR DEL GOLFO, S.C.'"
			End Select
			
			SQL = SQL & "SELECT 'IMPORTACION' as 'Tipo de Operacion', " & chr(13) & chr(10)
				SQL = SQL & "CASE i.cvecli01 " & chr(13) & chr(10)
				SQL = SQL & "WHEN '11000' THEN 'Virginia Leon' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '11001' THEN 'Gilberto Cruz' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '11002' THEN 'Iray Hinojosa' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '11003' THEN 'Francisco Bernal' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '14000' THEN 'Francisco Bernal' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '12000' THEN 'Georgina Perez' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '12001' THEN 'Lucero Bahena' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '12001' THEN 'Roberto Navarrete' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '12002' THEN 'Jorge Islas' " & chr(13) & chr(10)
				SQL = SQL & "WHEN '11004' THEN 'Monserrat Rodriguez' " & chr(13) & chr(10)
				SQL = SQL & "ELSE '' END AS 'Key Account', " & chr(13) & chr(10)
				SQL = SQL & "DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%m-%y') AS 'Fecha de Despacho', " & chr(13) & chr(10)
				SQL = SQL & "c.fdsp01, " & chr(13) & chr(10)
				SQL = SQL & "i.refcia01 AS 'Referencia', " & chr(13) & chr(10)
				SQL = SQL & "i.nomcli01 AS 'nomcli', " & chr(13) & chr(10)
				SQL = SQL & "i.cvecli01 AS 'cvecli', " & chr(13) & chr(10)
				SQL = SQL & "etx.m_observ AS 'Observaciones', " & chr(13) & chr(10)
				SQL = SQL & "cta.cgas31 AS 'Cuenta de Gastos', " & chr(13) & chr(10)
				SQL = SQL & "cta.fech31, " & chr(13) & chr(10)
				SQL = SQL & "cta.fech31 AS 'Fecha CG', " & chr(13) & chr(10)
				SQL = SQL & "cta.frec31 AS 'Fecha de Recepcion', " & chr(13) & chr(10)
				SQL = SQL & "cta.esta31 as 'Estado', " & chr(13) & chr(10)
				SQL = SQL & "DATE_FORMAT(i.fecpag01, '%y') as anopdto, " & chr(13) & chr(10)
				SQL = SQL & "i.cveadu01 as cveadu, " & chr(13) & chr(10)
				SQL = SQL & "i.patent01 as patente, " & chr(13) & chr(10)
				SQL = SQL & "i.numped01 as pdto, " & chr(13) & chr(10)
				SQL = SQL & adu & " as 'Oficina', " & chr(13) & chr(10)
				SQL = SQL & "i.nompro01 as 'Proveedor', " & chr(13) & chr(10)
				SQL = SQL & "fr.d_mer102 AS 'Descripcion de la Mercancia', " & chr(13) & chr(10)
				SQL = SQL & "if(cta.cgas31 is null or cta.cgas31 = '','Despachada pendiente de facturar',if(cta.frec31 is null or cta.frec31 ='' or cta.frec31 = '0000-00-00','Facturada pendiente de entregar','OK')) as 'Status', "  & chr(13) & chr(10)
				SQL = SQL & "IF(cau.c01causa is not null and cau.c01causa <> '', cau.c01causa, if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', if(cta.cgas31 is null and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada'))) as 'Causal', " & chr(13) & chr(10)
				SQL = SQL & "IF(cau.c01causa is not null and cau.c01causa <> '', cre.nom_resp, if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', if(cta.cgas31 is null and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado'))) as 'Responsable', " & chr(13) & chr(10)
				SQL = SQL & "ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', " & chr(13) & chr(10)
				SQL = SQL & "(TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos', " & chr(13) & chr(10)
				SQL = SQL & "IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion', " & chr(13) & chr(10)
				SQL = SQL & "d11.mont11 " & chr(13) & chr(10)
			SQL = SQL & "FROM "&aduanaTmp&"_extranet.ssdagi01 AS i " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.c01refer AS c ON i.refcia01 = c.refe01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.Numpat01 = i.patent01 AND bs.Detsit01 in (730,710) " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_status.c01caus as cau on cau.c01clavec = etx.clavec and (cau.c01tipoc = 'A' and cau.c01tipoo = '0') " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_status.cat_resp as cre on cre.id_resp = etx.id_resp " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 and (cta.frec31 = '0000-00-00' or cta.frec31 is null) and cta.esta31 <> 'C' " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31 and dl11.conc11 = 'LIQ' " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " & chr(13) & chr(10)
			SQL = SQL & "WHERE i.cveped01 <> 'R1' " & chr(13) & chr(10)
				SQL = SQL & "and (cta.fech31 is not null  or d11.conc11 is not null) " & chr(13) & chr(10)
				SQL = SQL & "and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " & chr(13) & chr(10)
				SQL = SQL & "and (bs.Fechst01 <> '0000-00-00' or bs.Fechst01 <> '' or bs.Fechst01 IS NOT NULL) " & chr(13) & chr(10)
				SQL = SQL & "and dl11.conc11 is null " & chr(13) & chr(10)
				SQL = SQL & "and (i.fecpag01 >= '"&DateI&"' and i.fecpag01 <= '"&DateF&"' ) " & chr(13) & chr(10)
				SQL = SQL & condicion & chr(13) & chr(10)
			SQL = SQL & "GROUP BY i.refcia01,cau.c01causa " & chr(13) & chr(10)
			SQL = SQL & "HAVING Estado is null or Estado = 'I' " & chr(13) & chr(10)
	end if
	
	  Response.Write(SQL)
	  Response.End
	GeneraSQL = SQL
end function

Function totalconte(refe)
	sqltotal = 	"SELECT COUNT(refe01) AS total " &_
				"FROM " & strOficina & "_extranet.d01conte " &_
				"WHERE refe01 = '" & refe & "'"
	Set RStotal = CreateObject("ADODB.Recordset")
	Set RStotal = ConnStr.Execute(sqltotal)
	IF RStotal.EOF = True And RStotal.BOF = True Then
		tot = 0
	Else
		RStotal.MoveFirst
		Do Until RStotal.EOF
			tot = RStotal.Fields.Item("total").Value
			RStotal.MoveNext
		Loop
	End If
	Rstotal.Close()
	Set RStotal = Nothing
	totalconte = tot
End Function

Function ImporteAnt(refe)
	SQLimp = 	""
	SQLimp = 	"SELECT refe11, " &_
				"DATE_FORMAT(MAX(fech11), '%d-%m-%Y') AS 'fecha', " &_
				"conc11, " &_
				"SUM(IF(conc11 = 'CAN', mont11*-1, mont11)) AS 'monto' " &_
				"FROM " & strOficina & "_extranet.d11movim " &_
				"WHERE (conc11 = 'ANT' OR conc11 = 'CAN') AND refe11 = '" & refe & "' " &_
				"GROUP BY refe11 "
	Set RSimp = Server.CreateObject("ADODB.Recordset")
	Set RSimp = ConnStr.Execute(SQLimp)
	If RSimp.BOF = True And RSimp.EOF = True Then
		import = 0
	Else
		import = RSimp.Fields.Item("monto").Value
	End If
	RSimp.Close()
	Set RSimp = Nothing
	' Response.Write(SQLimp)
	' Response.End()
	ImporteAnt = import
End Function

Function contienefacturas(refe)
	sqlfact = 	"SELECT i.refcia01, " &_
				"f.numfac39 " &_
				"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
				"INNER JOIN " & strOficina & "_extranet.ssfact39 AS f ON i.refcia01 = f.refcia39 " &_
				"AND i.patent01 = f.patent39 " &_
				"AND i.adusec01 = f.adusec39 " &_
				"WHERE i.refcia01 = '" & refe & "' "
	fact = ""
	Set RSfact = CreateObject("ADODB.RecordSet")
	Set RSfact = ConnStr.Execute(sqlfact)
	IF RSfact.EOF = True And RSfact.BOF = True Then
		fact = ""
	Else
		RSfact.MoveFirst
		Do Until RSfact.EOF
			fact = fact & RSfact.Fields.Item("numfac39").Value & ", "
			RSfact.MoveNext
		Loop
		fact = MID(fact,1,LEN(fact)-2)
	End If
	RSfact.Close()
	Set RSfact = Nothing
	contienefacturas = fact
end function

Function destinos(refe)
	SQL = ""
	desti = ""
	SQL = 	"SELECT count(DISTINCT(d01.marc01)) AS 'cuenta', " &_
			"d01.REFE01 AS referencia, " &_
			"d01.cdes01, " &_
			"c07.nomb07, " &_
			"d01.marc01 " &_
			"FROM rku_extranet.d01conte AS d01 " &_
			"LEFT JOIN " & stroficina & "_extranet.c01refer AS c01 ON c01.refe01 = d01.refe01 " &_
			"LEFT JOIN " & stroficina & "_extranet.c07desti AS c07 ON c07.cdes07 = d01.cdes01 " &_
			"WHERE d01.refe01 = '" & Refe & "' " &_
			"GROUP BY cdes01 "
	Set RSdest = CreateObject("ADODB.RecordSet")
	Set RSdest = ConnStr.Execute(SQL)
	IF RSdest.EOF = True And RSdest.BOF = True Then
		desti = ""
	Else
		RSdest.MoveFirst
		Do Until RSdest.EOF
			desti = desti & RSdest.Fields.Item("cuenta").Value & " " & RSdest.Fields.Item("nomb07").Value & ", "
			RSdest.MoveNext
		Loop
		desti = MID(desti,1,LEN(desti)-2)
	End If
	' Response.Write(desti)
	' Response.Write(SQL)
	' Response.End
	RSdest.Close()
	Set RSdest = Nothing
	destinos = desti
end function

Function Observaciones(refe)
	SQLObser = 	""
	observa = ""
	SQLObser = 	"SELECT c_referencia, " &_
				"REPLACE(m_observ,' ','&nbsp;') AS 'obser' " &_
				"FROM rku_status.etxpd " &_
				"WHERE c_referencia = '" & refe & "' and (clavec <> 0 or m_observ <> '') "
	Set RSobser = Server.CreateObject("ADODB.Recordset")
	Set RSobser = ConnStr.Execute(SQLObser)
	If RSobser.BOF = True And RSObser.EOF = True Then
		observa = ""
	Else
		RSobser.MoveFirst()
		Do Until RSobser.EOF = True
			observa = observa & RSobser.Fields.Item("obser").Value & " "
			RSobser.MoveNext()
		Loop
	End If
	RSobser.Close()
	Set RSobser = Nothing
	Observaciones = observa
End Function


Function Observaciones(refe)
	SQLobse = 	""
	observa =	""
	SQLobse = 	"SELECT DISTINCT c_referencia, " &_
				"REPLACE(REPLACE(m_observ, '\r', ''), '\n', '') AS m_observ " &_
				"FROM " & strOficina & "_status.etxpd " &_
				"WHERE c_referencia like '" & refe & "' " &_
				"AND m_observ IS NOT NULL " &_
				"AND TRIM(REPLACE(REPLACE(REPLACE(REPLACE(m_observ, ',', ''), '.', ''), '\r', ''), '\n', '')) <> '' " &_
				"AND m_observ NOT LIKE 'FECHA%IMPORTE%' "
	' Response.Write(SQLobse)
	' Response.End()
	Set RSobs = Server.CreateObject("ADODB.Recordset")
	Set RSobs = ConnStr.Execute(SQLobse)
	IF RSobs.BOF = True And RSobs.EOF = True Then
		observa = ""
	Else
		IF  RSobs.EOF = False then
			observa = RSobs("m_observ")
			RSobs.MoveNext()
			Do Until RSobs.EOF = True
				observa = observa & " | " &  RSobs("m_observ") 
				RSobs.MoveNext()
			Loop
		Else
			observa = RSobs("m_observ")
		end if
	End If
	RSobs.Close()
	Set RSobs = Nothing
	Observaciones = observa
End Function

Function Causales(refe, tipo)
	causas =	""
	SQLCausales = 	""
	SQLCausales = 	"SELECT DISTINCT etx.c_referencia, cau.c01causa, cau.c01tipoc " &_
					"FROM rku_status.etxpd AS etx " &_
					"INNER JOIN rku_status.c01caus AS cau ON cau.c01clavec = etx.clavec " &_
					"WHERE etx.c_referencia = '" & refe & "' AND cau.c01causa <> '' AND cau.c01tipoc LIKE '" & tipo & "'; "
	' if refe = "RKU10-08425" and tipo = "A" then
		' Response.Write(SQLCausales)
		' Response.End
	' end if
	Set RSCausas = Server.CreateObject("ADODB.RecordSet")
	Set RSCausas = ConnStr.Execute(SQLCausales)
	If RSCausas.BOF = True AND RSCausas.EOF = True Then
		causas = 	""
	Else
		RSCausas.MoveFirst()
		Do Until RSCausas.EOF = True
			Causas = Causas & RSCausas.Fields.Item("c01causa").Value & ", "
			RSCausas.MoveNext()
			' if RSCausas.EOF = True Then
				' Causas = Causas & RSCausas.Fields.Item("c01causa").Value
			' End If
		Loop
	End If
	' Response.Write(SQLCausales)
	' Response.end
	Causales = causas
End Function


Function KPI(opera, finicio, ffinal)
	SQL = 	""
	SQL = 	opera & "(IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'ALC', " &_
			"(( TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) ) -   " &_
			"if( ((DAYOFWEEK( " & finicio & " ) -1) = 6 )   , " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *1.5) - 0.5,  " &_
			"if( (DAYOFWEEK( " & finicio & " ) -1) = 7  ,   " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *1.5) - 1,  " &_
			"if(  ( (DAYOFWEEK( " & finicio & " ) -1)+TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) )  = 6, 0.5, " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *1.5) ))) " &_
			" - if( ((DAYOFWEEK( " & finicio & " ) -1) = 5 ), 0.5, 0)), " &_
			"(( TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) ) -   " &_
			"if( ((DAYOFWEEK( " & finicio & " ) -1) = 6 )   , " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *2) - 1,  " &_
			"if( (DAYOFWEEK( " & finicio & " ) -1) = 7  ,   " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *2) - 1,  " &_
			"if(  ( (DAYOFWEEK( " & finicio & " ) -1)+TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) )  = 6, 1, " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) * 2) ))) " &_
			" - if( ((DAYOFWEEK( " & finicio & " ) -1) = 5 ),1, 0) " &_
			" - if( ((DAYOFWEEK(" & ffinal & ") ) = 7 ),1, 0)))) "
			' Response.Write(SQL)
			' Response.End
	KPI = SQL
End Function


%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE SEGUIMIENTO DE OPERACIONES UNILEVER .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
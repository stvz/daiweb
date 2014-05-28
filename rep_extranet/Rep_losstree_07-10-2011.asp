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
	mov=request.form("mov")
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
	nocolumns = 12
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
								celdahead("Key Account") &_
								celdahead("Fecha de Despacho") &_
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Oficina") &_
								celdahead("Proveedor") &_
								celdahead("Descripcion de la Mercancia") &_
								celdahead("Causal") &_
								celdahead("Responsable") &_
								celdahead("Dias Transcurridos") & _
								celdahead("Importe Anticipo.")
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
								celdadatos(RSops.Fields.Item("Key Account").Value) &_
								celdadatos(RSops.Fields.Item("Fecha de Despacho").Value) &_
								celdadatos(RSops.Fields.Item("Referencia").Value) &_
								celdadatos(Cstr(RSops.Fields.Item("anopdto").Value) & " " & Cstr(RSops.Fields.Item("cveadu").Value) & " " & Cstr(RSops.Fields.Item("patente").Value) & " " & Cstr(RSops.Fields.Item("pdto").Value)) &_
								celdadatos(RSops.Fields.Item("Oficina").Value) &_
								celdadatos(RSops.Fields.Item("Proveedor").Value) &_
								celdadatos(RSops.Fields.Item("Descripcion de la Mercancia").Value) &_
								celdadatos(RSops.Fields.Item("Causal").Value) &_
								celdadatos(RSops.Fields.Item("Responsable").Value) &_
								celdadatos(RSops.Fields.Item("Dias Transcurridos").Value) & _
								celdadatos(RSops.Fields.Item("mont11").Value)
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
			condicion = "AND i.rfccli01 = 'UME651115N48' "
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
		
	SQL = 	"SELECT 'IMPORTACION' as 'Tipo de Operacion', " &_
					"CASE i.cvecli01 " &_
						"WHEN '11000' THEN 'Virginia Leon' " &_
						"WHEN '11001' THEN 'Gilberto Cruz' " &_
						"WHEN '11002' THEN 'Iray Hinojosa' " &_
						"WHEN '11003' THEN 'Francisco Bernal' " &_
						"WHEN '14000' THEN 'Francisco Bernal' " & _
						" WHEN '12000' THEN 'Georgina Perez'  " & _
						" WHEN '12001' THEN 'Lucero Bahena'  " & _
						" WHEN '12001' THEN 'Roberto Navarrete'  " & _
						" WHEN '12002' THEN 'Jorge Islas'  " & _
						"WHEN '11004' THEN 'Monserrat Rodriguez' " &_
					"ELSE '' END AS 'Key Account'," &_
					"DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%b-%y') AS 'Fecha de Despacho', " &_
					"i.refcia01 AS 'Referencia', " &_
					"DATE_FORMAT(i.fecpag01, '%y') as anopdto, " &_
					"i.cveadu01 as cveadu, " &_
					"i.patent01 as patente, " &_
					"i.numped01 as pdto, " &_
					"'10110825 - GRUPO REYES KURI,  S.C.' as 'Oficina', " &_
					"i.nompro01 as 'Proveedor', " &_
					"fr.d_mer102 AS 'Descripcion de la Mercancia', " &_
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   " & _
					"  cau.c01causa, " & _
					"     if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', " & _
					"      if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada')))  as 'Causal',  " & _
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   "  & _
					" cre.nom_resp, "  & _
					"    if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', "  & _
					"     if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado')))  as 'Responsable',  "  & _
					
					"ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx',  (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos'," &_
					"IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion',d11.mont11 " &_
			"FROM rku_extranet.ssdagi01 AS i " &_
					"LEFT JOIN rku_extranet.c01refer AS c ON i.refcia01 = c.refe01 and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " &_
					"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 AND (bs.Detsit01 = 730 or bs.Detsit01 = 710) " &_
					"LEFT JOIN rku_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " &_
					"LEFT JOIN rku_status.c01caus as cau on cau.c01clavec = etx.clavec and cau.c01tipoc = 'A' and cau.c01tipoo = '0' " &_
					"LEFT JOIN rku_status.cat_resp as cre on cre.id_resp = etx.id_resp " &_
					"LEFT JOIN rku_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
					"LEFT JOIN rku_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
					"LEFT JOIN rku_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN rku_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND (cta.esta31 <> 'C' or cta.esta31 IS NOT NULL) " &_
					"LEFT JOIN rku_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & _
					" LEFT JOIN rku_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31  and dl11.conc11 = 'LIQ' " & _
			"WHERE " & _
				" i.cveped01 <> 'R1'  " & _
				" and (cta.frec31 = '0000-00-00'  or cta.frec31 is null  )" & _
				" and (bs.Fechst01 is not null or cta.frec31 is not null or cta.fech31 is not null  or d11.conc11 is not null)" & _
				" and dl11.conc11 is null " & _
				" and (i.fecpag01 >= '" & DateI & "'" & _
				" and i.fecpag01 <=  '"  & DateF & "' ) "  & condicion &_
			"GROUP BY i.refcia01, cau.c01causa " &_
			
			"UNION ALL " &_
		
			"SELECT 'IMPORTACION' as 'Tipo de Operacion', " &_
					"CASE i.cvecli01 " &_
						"WHEN '11000' THEN 'Virginia Leon' " &_
						"WHEN '11001' THEN 'Gilberto Cruz' " &_
						"WHEN '11002' THEN 'Iray Hinojosa' " &_
						"WHEN '11003' THEN 'Francisco Bernal' " &_
						"WHEN '14000' THEN 'Francisco Bernal' " & _
						" WHEN '12000' THEN 'Georgina Perez'  " & _
						" WHEN '12001' THEN 'Lucero Bahena'  " & _
						" WHEN '12001' THEN 'Roberto Navarrete'  " & _
						" WHEN '12002' THEN 'Jorge Islas'  " & _
						"WHEN '11004' THEN 'Monserrat Rodriguez' " &_
					"ELSE '' END AS 'Key Account'," &_
					"DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%b-%y') AS 'Fecha de Despacho', " &_
					"i.refcia01 AS 'Referencia', " &_
					"DATE_FORMAT(i.fecpag01, '%y') as anopdto, " &_
					"i.cveadu01 as cveadu, " &_
					"i.patent01 as patente, " &_
					"i.numped01 as pdto, " &_
					"'10110818 - DESPACHOS AEREOS INTEGRADOS, S.C.' as 'Oficina', " &_
					"i.nompro01 as 'Proveedor', " &_
					"fr.d_mer102 AS 'Descripcion de la Mercancia', " &_
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   " & _
					"  cau.c01causa, " & _
					"     if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', " & _
					"      if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada')))  as 'Causal',  " & _
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   "  & _
					" cre.nom_resp, "  & _
					"    if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', "  & _
					"     if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado')))  as 'Responsable',  "  & _
					
				"ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos'," &_
					"IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion',d11.mont11 " &_
			"FROM dai_extranet.ssdagi01 AS i " &_
					"LEFT JOIN dai_extranet.c01refer AS c ON i.refcia01 = c.refe01 and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " &_
					"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 AND (bs.Detsit01 = 730 or bs.Detsit01 = 710) " &_
					"LEFT JOIN dai_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " &_
					"LEFT JOIN dai_status.c01caus as cau on cau.c01clavec = etx.clavec and cau.c01tipoc = 'A' and cau.c01tipoo = '0' " &_
					"LEFT JOIN dai_status.cat_resp as cre on cre.id_resp = etx.id_resp " &_
					"LEFT JOIN dai_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
					"LEFT JOIN dai_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
					"LEFT JOIN dai_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN dai_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND (cta.esta31 <> 'C' or cta.esta31 IS NOT NULL) " &_
					"LEFT JOIN dai_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & _
										" LEFT JOIN dai_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31  and dl11.conc11 = 'LIQ' " & _
			"WHERE " & _
				" i.cveped01 <> 'R1'  " & _
				" and (cta.frec31 = '0000-00-00'  or cta.frec31 is null  )" & _
				" and (bs.Fechst01 is not null or cta.frec31 is not null or cta.fech31 is not null  or d11.conc11 is not null)" & _
				" and dl11.conc11 is null " & _
				" and (i.fecpag01 >= '" & DateI & "'" & _
				" and i.fecpag01 <=  '"  & DateF & "' ) "  & condicion &_
			"GROUP BY i.refcia01, cau.c01causa " &_

			"UNION ALL " &_

			"SELECT 'IMPORTACION' as 'Tipo de Operacion', " &_
					"CASE i.cvecli01 " &_
						"WHEN '11000' THEN 'Virginia Leon' " &_
						"WHEN '11001' THEN 'Gilberto Cruz' " &_
						"WHEN '11002' THEN 'Iray Hinojosa' " &_
						"WHEN '11003' THEN 'Francisco Bernal' " &_
						"WHEN '14000' THEN 'Francisco Bernal' " & _
						" WHEN '12000' THEN 'Georgina Perez'  " & _
						" WHEN '12001' THEN 'Lucero Bahena'  " & _
						" WHEN '12001' THEN 'Roberto Navarrete'  " & _
						" WHEN '12002' THEN 'Jorge Islas'  " & _
						"WHEN '11004' THEN 'Monserrat Rodriguez' " &_
					"ELSE '' END AS 'Key Account'," &_
					"DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%b-%y') AS 'Fecha de Despacho', " &_
					"i.refcia01 AS 'Referencia', " &_
					"DATE_FORMAT(i.fecpag01, '%y') as anopdto, " &_
					"i.cveadu01 as cveadu, " &_
					"i.patent01 as patente, " &_
					"i.numped01 as pdto, " &_
					"'10080746 - SERVADUANALES DEL PACIFICO, S.C.' as 'Oficina', " &_
					"i.nompro01 as 'Proveedor', " &_
					"fr.d_mer102 AS 'Descripcion de la Mercancia', " &_
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   " & _
					"  cau.c01causa, " & _
					"     if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', " & _
					"      if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada')))  as 'Causal',  " & _
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   "  & _
					" cre.nom_resp, "  & _
					"    if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', "  & _
					"     if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado')))  as 'Responsable',  "  & _
					
					"ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos'," &_
					"IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion',d11.mont11 " &_
			"FROM sap_extranet.ssdagi01 AS i " &_
					"LEFT JOIN sap_extranet.c01refer AS c ON i.refcia01 = c.refe01 and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " &_
					"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 AND (bs.Detsit01 = 730 or bs.Detsit01 = 710) " &_
					"LEFT JOIN sap_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " &_
					"LEFT JOIN sap_status.c01caus as cau on cau.c01clavec = etx.clavec and cau.c01tipoc = 'A' and cau.c01tipoo = '0' " &_
					"LEFT JOIN sap_status.cat_resp as cre on cre.id_resp = etx.id_resp " &_
					"LEFT JOIN sap_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
					"LEFT JOIN sap_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
					"LEFT JOIN sap_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN sap_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND (cta.esta31 <> 'C' or cta.esta31 IS NOT NULL) " &_
					"LEFT JOIN sap_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & _
										" LEFT JOIN sap_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31  and dl11.conc11 = 'LIQ' " & _
			"WHERE " & _
				" i.cveped01 <> 'R1'  " & _
				" and (cta.frec31 = '0000-00-00'  or cta.frec31 is null  )" & _
				" and (bs.Fechst01 is not null or cta.frec31 is not null or cta.fech31 is not null  or d11.conc11 is not null)" & _
				" and dl11.conc11 is null " & _
				" and (i.fecpag01 >= '" & DateI & "'" & _
				" and i.fecpag01 <=  '"  & DateF & "' ) "  & condicion &_
			"GROUP BY i.refcia01, cau.c01causa " &_

			"UNION ALL " &_

			"SELECT 'IMPORTACION' as 'Tipo de Operacion', " &_
					"CASE i.cvecli01 " &_
						"WHEN '11000' THEN 'Virginia Leon' " &_
						"WHEN '11001' THEN 'Gilberto Cruz' " &_
						"WHEN '11002' THEN 'Iray Hinojosa' " &_
						"WHEN '11003' THEN 'Francisco Bernal' " &_
						"WHEN '14000' THEN 'Francisco Bernal' " & _
						" WHEN '12000' THEN 'Georgina Perez'  " & _
						" WHEN '12001' THEN 'Lucero Bahena'  " & _
						" WHEN '12001' THEN 'Roberto Navarrete'  " & _
						" WHEN '12002' THEN 'Jorge Islas'  " & _
						"WHEN '11004' THEN 'Monserrat Rodriguez' " &_
					"ELSE '' END AS 'Key Account'," &_
					"DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%b-%y') AS 'Fecha de Despacho', " &_
					"i.refcia01 AS 'Referencia', " &_
					"DATE_FORMAT(i.fecpag01, '%y') as anopdto, " &_
					"i.cveadu01 as cveadu, " &_
					"i.patent01 as patente, " &_
					"i.numped01 as pdto, " &_
					"'10110819 - COMERCIO EXTERIOR DEL GOLFO, S.C.' as 'Oficina', " &_
					"i.nompro01 as 'Proveedor', " &_
					"fr.d_mer102 AS 'Descripcion de la Mercancia', " &_

					" IF(cau.c01causa is not null and cau.c01causa <> '',   " & _
					"  cau.c01causa, " & _
					"     if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', " & _
					"      if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada')))  as 'Causal',  " & _
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   "  & _
					" cre.nom_resp, "  & _
					"    if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', "  & _
					"     if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado')))  as 'Responsable',  "  & _
					
					"ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos'," &_
					"IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion' ,d11.mont11 " &_
			"FROM ceg_extranet.ssdagi01 AS i " &_
					"LEFT JOIN ceg_extranet.c01refer AS c ON i.refcia01 = c.refe01 and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " &_
					"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 AND (bs.Detsit01 = 730 or bs.Detsit01 = 710) " &_
					"LEFT JOIN ceg_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " &_
					"LEFT JOIN ceg_status.c01caus as cau on cau.c01clavec = etx.clavec and cau.c01tipoc = 'A' and cau.c01tipoo = '0' " &_
					"LEFT JOIN ceg_status.cat_resp as cre on cre.id_resp = etx.id_resp " &_
					"LEFT JOIN ceg_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
					"LEFT JOIN ceg_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
					"LEFT JOIN ceg_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN ceg_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND (cta.esta31 <> 'C' or cta.esta31 IS NOT NULL) " &_
					"LEFT JOIN ceg_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & _
										" LEFT JOIN ceg_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31  and dl11.conc11 = 'LIQ' " & _
			"WHERE " & _
				" i.cveped01 <> 'R1'  " & _
				" and (cta.frec31 = '0000-00-00'  or cta.frec31 is null  )" & _
				" and (bs.Fechst01 is not null or cta.frec31 is not null or cta.fech31 is not null  or d11.conc11 is not null)" & _
				" and dl11.conc11 is null " & _
				" and (i.fecpag01 >= '" & DateI & "'" & _
				" and i.fecpag01 <=  '"  & DateF & "' ) "  & condicion &_
			"GROUP BY i.refcia01, cau.c01causa " &_

			"UNION ALL " &_

			"SELECT 'IMPORTACION' as 'Tipo de Operacion', " &_
					"CASE i.cvecli01 " &_
						"WHEN '11000' THEN 'Virginia Leon' " &_
						"WHEN '11001' THEN 'Gilberto Cruz' " &_
						"WHEN '11002' THEN 'Iray Hinojosa' " &_
						"WHEN '11003' THEN 'Francisco Bernal' " &_
						"WHEN '14000' THEN 'Francisco Bernal' " & _
						" WHEN '12000' THEN 'Georgina Perez'  " & _
						" WHEN '12001' THEN 'Lucero Bahena'  " & _
						" WHEN '12001' THEN 'Roberto Navarrete'  " & _
						" WHEN '12002' THEN 'Jorge Islas'  " & _
						"WHEN '11004' THEN 'Monserrat Rodriguez' " &_
					"ELSE '' END AS 'Key Account'," &_
					"DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%b-%y') AS 'Fecha de Despacho', " &_
					"i.refcia01 AS 'Referencia', " &_
					"DATE_FORMAT(i.fecpag01, '%y') as anopdto, " &_
					"i.cveadu01 as cveadu, " &_
					"i.patent01 as patente, " &_
					"i.numped01 as pdto, " &_
					"'10080746 - SERVADUANALES DEL PACIFICO, S.C.' as 'Oficina', " &_
					"i.nompro01 as 'Proveedor', " &_
					"fr.d_mer102 AS 'Descripcion de la Mercancia', " &_

					" IF(cau.c01causa is not null and cau.c01causa <> '',   " & _
					"  cau.c01causa, " & _
					"     if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', " & _
					"      if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada')))  as 'Causal',  " & _
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   "  & _
					" cre.nom_resp, "  & _
					"    if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', "  & _
					"     if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado')))  as 'Responsable',  "  & _
					
				"ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos'," &_
					"IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion' ,d11.mont11 " &_
			"FROM lzr_extranet.ssdagi01 AS i " &_
					"LEFT JOIN lzr_extranet.c01refer AS c ON i.refcia01 = c.refe01 and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " &_
					"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 AND (bs.Detsit01 = 730 or bs.Detsit01 = 710) " &_
					"LEFT JOIN lzr_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " &_
					"LEFT JOIN lzr_status.c01caus as cau on cau.c01clavec = etx.clavec and cau.c01tipoc = 'A' and cau.c01tipoo = '0' " &_
					"LEFT JOIN lzr_status.cat_resp as cre on cre.id_resp = etx.id_resp " &_
					"LEFT JOIN lzr_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
					"LEFT JOIN lzr_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
					"LEFT JOIN lzr_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN lzr_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND (cta.esta31 <> 'C' or cta.esta31 IS NOT NULL) " &_
					"LEFT JOIN lzr_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & _
										" LEFT JOIN lzr_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31  and dl11.conc11 = 'LIQ' " & _
			"WHERE " & _
				" i.cveped01 <> 'R1'  " & _
				" and (cta.frec31 = '0000-00-00'  or cta.frec31 is null  )" & _
				" and (bs.Fechst01 is not null or cta.frec31 is not null or cta.fech31 is not null  or d11.conc11 is not null)" & _
				" and dl11.conc11 is null " & _
				" and (i.fecpag01 >= '" & DateI & "'" & _
				" and i.fecpag01 <=  '"  & DateF & "' ) "  & condicion &_
			"GROUP BY i.refcia01, cau.c01causa " &_

			"UNION ALL " &_

			"SELECT 'IMPORTACION' as 'Tipo de Operacion', " &_
					"CASE i.cvecli01 " &_
						"WHEN '11000' THEN 'Virginia Leon' " &_
						"WHEN '11001' THEN 'Gilberto Cruz' " &_
						"WHEN '11002' THEN 'Iray Hinojosa' " &_
						"WHEN '11003' THEN 'Francisco Bernal' " &_
						"WHEN '14000' THEN 'Francisco Bernal' " & _
						" WHEN '12000' THEN 'Georgina Perez'  " & _
						" WHEN '12001' THEN 'Lucero Bahena'  " & _
						" WHEN '12001' THEN 'Roberto Navarrete'  " & _
						" WHEN '12002' THEN 'Jorge Islas'  " & _
						"WHEN '11004' THEN 'Monserrat Rodriguez' " &_
					"ELSE '' END AS 'Key Account'," &_
					"DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%b-%y') AS 'Fecha de Despacho', " &_
					"i.refcia01 AS 'Referencia', " &_
					"DATE_FORMAT(i.fecpag01, '%y') as anopdto, " &_
					"i.cveadu01 as cveadu, " &_
					"i.patent01 as patente, " &_
					"i.numped01 as pdto, " &_
					"'10110819 - COMERCIO EXTERIOR DEL GOLFO, S.C.' as 'Oficina', " &_
					"i.nompro01 as 'Proveedor', " &_
					"fr.d_mer102 AS 'Descripcion de la Mercancia', " &_

					" IF(cau.c01causa is not null and cau.c01causa <> '',   " & _
					"  cau.c01causa, " & _
					"     if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'2.6 FACTURADA PENDIENTE DE INGRESAR', " & _
					"      if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'2.3 EN TIEMPO. ANTICIPO RECIBIDO MENOR A 15 DIAS','No Capturada')))  as 'Causal',  " & _
					
					" IF(cau.c01causa is not null and cau.c01causa <> '',   "  & _
					" cre.nom_resp, "  & _
					"    if((cta.frec31 is null or cta.frec31 = '0000-00-00') and cta.cgas31 is not null ,'GRUPO ZEGO', "  & _
					"     if(cta.cgas31 is null  and (TO_DAYS( sysdate() ) - TO_DAYS(d11.fech11) ) <15,'GRUPO ZEGO','No Capturado')))  as 'Responsable',  "  & _
					
					"ABS(IF( CURDATE() >= DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY) , DATEDIFF(CURDATE(),DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY)) , DATEDIFF(DATE_ADD(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),INTERVAL 5 DAY),CURDATE()))) as 'Dias Transcurridosx', (TO_DAYS( sysdate() ) - TO_DAYS(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) )) as 'Dias Transcurridos'," &_
					"IF(etx.f_solucion is not null and etx.f_solucion <> '' and etx.f_solucion <> '0000-00-00',DATE_FORMAT(etx.f_solucion,'%d-%b-%y'),'No Capturada') as 'Fecha Solucion',d11.mont11 " &_
			"FROM tol_extranet.ssdagi01 AS i " &_
					"LEFT JOIN tol_extranet.c01refer AS c ON i.refcia01 = c.refe01 and (c.fdsp01 <> '0000-00-00' or c.fdsp01 <> '' or c.fdsp01 IS NOT NULL) " &_
					"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 AND (bs.Detsit01 = 730 or bs.Detsit01 = 710) " &_
					"LEFT JOIN tol_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 " &_
					"LEFT JOIN tol_status.c01caus as cau on cau.c01clavec = etx.clavec and cau.c01tipoc = 'A' and cau.c01tipoo = '0' " &_
					"LEFT JOIN tol_status.cat_resp as cre on cre.id_resp = etx.id_resp " &_
					"LEFT JOIN tol_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
					"LEFT JOIN tol_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
					"LEFT JOIN tol_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN tol_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND (cta.esta31 <> 'C' or cta.esta31 IS NOT NULL) " &_
					"LEFT JOIN tol_extranet.d11movim as d11 on d11.refe11= i.refcia01 and d11.conc11 = 'ANT' " & _
										" LEFT JOIN tol_extranet.d11movim as dl11 on dl11.refe11= cta.cgas31  and dl11.conc11 = 'LIQ' " & _
			"WHERE " & _
				" i.cveped01 <> 'R1'  " & _
				" and (cta.frec31 = '0000-00-00'  or cta.frec31 is null  )" & _
				" and (bs.Fechst01 is not null or cta.frec31 is not null or cta.fech31 is not null  or d11.conc11 is not null)" & _
				" and dl11.conc11 is null " & _
				" and (i.fecpag01 >= '" & DateI & "'" & _
				" and i.fecpag01 <=  '"  & DateF & "' ) "  & condicion &_
			"GROUP BY i.refcia01, cau.c01causa " &_

			"ORDER BY 14 DESC "
	 'Response.Write(SQL)
	 ' Response.End
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
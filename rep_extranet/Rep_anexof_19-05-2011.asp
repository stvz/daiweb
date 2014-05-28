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

	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & Diai

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	nocolumns = 0
	tablamov = ""
	mov = "i"
	if mov = "i" then
		movi = ":: IMPORTACION ::"
		tablamov = "ssdagi01"
		nocolumns = 66
		query = GeneraSQL
	else
		movi = ":: EXPORTACION ::"
		tablamov = "ssdage01"
		nocolumns = 17
		query = GeneraSQL
	end if
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write("DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	' Response.Write(query)
	' Response.End()
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left"">" &_
												"::.... REPORTE ANEXO F .... ::" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
												movi &_ 
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
								celdahead("BU") &_
								celdahead("Agente Aduanal") &_
								celdahead("Nombre Aduana") &_
								celdahead("Aduana") &_
								celdahead("Seccion") &_
								celdahead("Fecha De Pago") &_
								celdahead("Fecha De Cruce") &_
								celdahead("Patente") &_
								celdahead("No. Pedimento") &_
								celdahead("Sec Fraccion") &_
								celdahead("Clave Pedimento") &_
								celdahead("Pedimento Original") &_
								celdahead("Fecha Pedimento Original") &_
								celdahead("IVA del Pedimento Original") &_
								celdahead("IGI del Pedimento Original") &_
								celdahead("Modal") &_
								celdahead("Referencia / Pedido") &_
								celdahead("Facturas") &_
								celdahead("Num Parte") &_
								celdahead("Proveedor") &_
								celdahead("Material (1)") &_
								celdahead("Fraccion") &_
								celdahead("Material (2)") &_
								celdahead("Valor Factura USD") &_
								celdahead("Tipo de Cambio") &_
								celdahead("Valor MN") &_
								celdahead("Valor Aduana") &_
								celdahead("Peso Bruto") &_
								celdahead("P_Advalorem") &_
								celdahead("Advalorem") &_
								celdahead("IVA") &_
								celdahead("DTA") &_
								celdahead("Maniobras") &_
								celdahead("PRV") &_
								celdahead("Validacion") &_
								celdahead("Serv_Complementarios") &_
								celdahead("Demoras") &_
								celdahead("ECI") &_
								celdahead("Otros Pagos") &_
								celdahead("Maniobras Puerto") &_
								celdahead("Honorarios") &_
								celdahead("Reconocimiento Aduanero") &_
								celdahead("IVA Honorarios") &_
								celdahead("Cuenta de Gastos") &_
								celdahead("Importe de Cuenta de Gastos") &_
								celdahead("Transportistas") &_
								celdahead("Flete Terrestre") &_
								celdahead("Naviera") &_
								celdahead("Flete Maritimo") &_
								celdahead("Linea Aerea") &_
								celdahead("Flete Aereo") &_
								celdahead("Pais Vendedor") &_
								celdahead("Pais Origen") &_
								celdahead("Beneficio Arancelario") &_
								celdahead("Incoterm") &_
								celdahead("Vinculacion") &_
								celdahead("Valoracion") &_
								celdahead("Incrementables") &_
								celdahead("Decrementables") &_
								celdahead("Rectificacion") &_
								celdahead("FECHA") &_
								celdahead("NUM_POLIZA") &_
								celdahead("TIPO_POLIZA") &_
								celdahead("IMPORTE_MXP") &_
								celdahead("CARPETA") &_
								celdahead("Contenedores") 
		
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
								celdadatos(RSops.Fields.Item("BU").Value) &_
								celdadatos(RSops.Fields.Item("Agente Aduanal").Value) &_
								celdadatos(RSops.Fields.Item("Nombre Aduana").Value) &_
								celdadatos(RSops.Fields.Item("Aduana").Value) &_
								celdadatos(RSops.Fields.Item("Seccion").Value) &_
								celdadatos(RSops.Fields.Item("Fecha de Pago").Value) &_
								celdadatos(RSops.Fields.Item("Fecha de Cruce").Value) &_
								celdadatos(RSops.Fields.Item("Patente").Value) &_
								celdadatos(RSops.Fields.Item("No. Pedimento").Value) &_
								celdadatos(RSops.Fields.Item("Sec Fraccion").Value) &_
								celdadatos(RSops.Fields.Item("Clave Pedimento").Value) &_
								celdadatos(RSops.Fields.Item("Pedimento Original").Value) &_
								celdadatos(RSops.Fields.Item("Fecha Pedimento Original").Value) &_
								celdadatos(RSops.Fields.Item("IVA del Pedimento Original").Value) &_
								celdadatos(RSops.Fields.Item("IGI del Pedimento Original").Value) &_
								celdadatos(RSops.Fields.Item("Modal").Value) &_
								celdadatos(RSops.Fields.Item("Referencia / Pedido").Value) &_
								celdadatos(RSops.Fields.Item("Facturas").Value) &_
								celdadatos(RSops.Fields.Item("Num Parte").Value) &_
								celdadatos(RSops.Fields.Item("Proveedor").Value) &_
								celdadatos(RSops.Fields.Item("Material (1)").Value) &_
								celdadatos(RSops.Fields.Item("Fraccion").Value) &_
								celdadatos(RSops.Fields.Item("Material (2)").Value) &_
								celdadatos(RSops.Fields.Item("Valor Factura USD").Value) &_
								celdadatos(RSops.Fields.Item("Tipo de Cambio").Value) &_
								celdadatos(RSops.Fields.Item("Valor MN").Value) &_
								celdadatos(RSops.Fields.Item("Valor Aduana").Value) &_
								celdadatos(RSops.Fields.Item("Peso Bruto").Value) &_
								celdadatos(RSops.Fields.Item("P_Advalorem").Value) &_
								celdadatos(RSops.Fields.Item("Advalorem").Value) &_
								celdadatos(RSops.Fields.Item("IVA").Value) &_
								celdadatos(RSops.Fields.Item("DTA").Value) &_
								celdadatos(RSops.Fields.Item("Maniobras").Value) &_
								celdadatos(RSops.Fields.Item("PRV").Value) &_
								celdadatos(RSops.Fields.Item("Validacion").Value) &_
								celdadatos(RSops.Fields.Item("Serv_Complementarios").Value) &_
								celdadatos(RSops.Fields.Item("Demoras").Value) &_
								celdadatos(RSops.Fields.Item("ECI").Value) &_
								celdadatos(RSops.Fields.Item("Otros Pagos").Value) &_
								celdadatos(RSops.Fields.Item("Maniobras Puerto").Value) &_
								celdadatos(RSops.Fields.Item("Honorarios").Value) &_
								celdadatos(RSops.Fields.Item("Reconocimiento Aduanero").Value) &_
								celdadatos(RSops.Fields.Item("IVA Honorarios").Value) &_
								celdadatos(RSops.Fields.Item("Cuenta de Gastos").Value) &_
								celdadatos(RSops.Fields.Item("Importe Cuenta de Gastos").Value) &_
								celdadatos(RSops.Fields.Item("transportista").Value) &_
								celdadatos(RSops.Fields.Item("Flete Terrestre").Value) &_
								celdadatos(RSops.Fields.Item("Naviera").Value) &_
								celdadatos(RSops.Fields.Item("Flete Maritimo").Value) &_
								celdadatos(RSops.Fields.Item("Linea Aerea").Value) &_
								celdadatos(RSops.Fields.Item("Flete aereo").Value) &_
								celdadatos(RSops.Fields.Item("Pais Vendedor").Value) &_
								celdadatos(RSops.Fields.Item("Pais Origen").Value) &_
								celdadatos(RSops.Fields.Item("Beneficio Arancelario").Value) &_
								celdadatos(RSops.Fields.Item("Incoterm").Value) &_
								celdadatos(RSops.Fields.Item("Vinculacion").Value) &_
								celdadatos(RSops.Fields.Item("Valoracion").Value) &_
								celdadatos(RSops.Fields.Item("Incrementables").Value) &_
								celdadatos(RSops.Fields.Item("Decrementables").Value) &_
								celdadatos(RSops.Fields.Item("Rectificacion").Value) &_
								celdadatos(RSops.Fields.Item("Fecha").Value) &_
								celdadatos(RSops.Fields.Item("NUM_POLIZA").Value) &_
								celdadatos(RSops.Fields.Item("TIPO_POLIZA").Value) &_
								celdadatos(RSops.Fields.Item("IMPORTE_MXP").Value) &_
								celdadatos(RSops.Fields.Item("CARPETA").Value) &_
								celdadatos(RSops.Fields.Item("Contenedores").Value)
								
								
			
			datos = datos &	"</tr>"
			Rsops.MoveNext()
		Loop
	

	' Response.Write(info & header & datos & "</table><br>" & prom)
	' Response.End()
	html = Actualizaciones & info & header & datos & "</table><br>"
	
	
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
		condicion = "AND i.rfccli01 = '" & Vrfc & "' "
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
	condicion = filtro
	SQL = 	"SELECT DISTINCT " &_
			"CASE ops.cvecli01 " &_
				"WHEN '4284' THEN '5636' " &_
				"ELSE '-------' " &_
			"END AS 'BU', " &_
			"con.agente32 AS 'Agente Aduanal', " &_
			"ops.nomadu01 AS 'Nombre Aduana', " &_
			"ops.cveadu01 AS 'Aduana', " &_
			"ops.cvesec01 AS 'Seccion', " &_
			"DATE_FORMAT(IF(ops.cveped01 = 'R1', rec.f_pagr06, ops.fecpag01), '%d/%m/%Y') AS 'Fecha de Pago', " &_
			"IFNULL(DATE_FORMAT((SELECT MIN(t.FecTick) " &_
			"FROM rku_cpsimples.tcepartidas AS t " &_
			"WHERE MID(t.Refe01, 1, 11) = MID(ops.refcia01, 1, 11) AND t.Frac01 = fra.fraarn02 " &_
			"GROUP BY MID(t.Refe01, 1, 11)), '%d-%m-%Y'), '') AS 'Fecha de Cruce', " &_
			"ops.patent01 AS 'Patente', " &_
			"ops.numped01 AS 'No. Pedimento', " &_
			"fra.ordfra02 AS 'Sec Fraccion', " &_
			"ops.cveped01 AS 'Clave Pedimento', " &_
			"CAST(CONCAT_WS(' ', rec.cveado06, rec.agente06, rec.pedorg06) AS CHAR) AS 'Pedimento Original', " &_
			"IF(ops.cveped01 = 'R1', DATE_FORMAT(ops.fecpag01, '%d/%m/%Y'), '') AS 'Fecha Pedimento Original', " &_
			"IF(ops.cveped01 = 'R1', " &_
			"IFNULL((SELECT SUM(iva.import36) " &_
			"FROM rku_extranet.sscont36 AS iva " &_
			"WHERE iva.refcia36 = rec.reforg06 AND iva.cveimp36 = 3 " &_
			"GROUP BY iva.refcia36 " &_
			"), 0) " &_
			", 0) AS 'IVA del Pedimento Original', " &_
			"IF(ops.cveped01 = 'R1', " &_
			"IFNULL((SELECT SUM(adv.import36) " &_
			"FROM rku_extranet.sscont36 AS adv " &_
			"WHERE adv.refcia36 = rec.reforg06 AND adv.cveimp36 = 6 " &_
			"GROUP BY adv.refcia36 " &_
			"), 0) " &_
			", 0) AS 'IGI del Pedimento Original', " &_
			"IF(ops.tipped01 = 1, 'Completo', 'Consolidado') AS 'Modal', " &_
			"(SELECT GROUP_CONCAT(DISTINCT(d05.pedi05)) " &_
			"FROM rku_extranet.d05artic AS d05 " &_
			"WHERE d05.refe05 = ops.refcia01 AND d05.agru05 = fra.ordfra02 AND d05.frac05 = fra.fraarn02 " &_
			"GROUP BY d05.refe05) AS 'Referencia / Pedido', " &_
			"(SELECT GROUP_CONCAT(DISTINCT(d05.fact05)) " &_
			"FROM rku_extranet.d05artic AS d05 " &_
			"WHERE d05.refe05 = ops.refcia01 AND d05.agru05 = fra.ordfra02 AND d05.frac05 = fra.fraarn02 " &_
			"GROUP BY d05.refe05) AS 'Facturas', " &_
			"(SELECT GROUP_CONCAT(DISTINCT(d05.cpro05)) " &_
			"FROM rku_extranet.d05artic AS d05 " &_
			"WHERE d05.refe05 = ops.refcia01 AND d05.agru05 = fra.ordfra02 AND d05.frac05 = fra.fraarn02 " &_
			"GROUP BY d05.refe05) AS 'Num Parte', " &_
			"(SELECT GROUP_CONCAT(DISTINCT(fac.nompro39)) " &_
			"FROM rku_extranet.d05artic AS d05 " &_
			"LEFT JOIN rku_extranet.ssfact39 AS fac ON fac.refcia39 = d05.refe05 AND fac.numfac39 = d05.fact05 " &_
			"WHERE d05.refe05 = ops.refcia01 AND d05.agru05 = fra.ordfra02 AND d05.frac05 = fra.fraarn02 AND fac.patent39 = ops.patent01 " &_
			"GROUP BY d05.refe05) AS 'Proveedor', " &_
			"fra.d_mer102 AS 'Material (1)', " &_
			"fra.fraarn02 AS 'Fraccion', " &_
			"fra.d_mer202 AS 'Material (2)', " &_
			"ROUND(fra.vmerme02 * fra.factmo02, 2) AS 'Valor Factura USD', " &_
			"ops.tipcam01 AS 'Tipo de Cambio', " &_
			"ROUND(fra.vmerme02 * fra.factmo02 * ops.tipcam01, 2) AS 'Valor MN', " &_
			"ROUND(fra.vaduan02, 0) AS 'Valor Aduana', " &_
			"ops.pesobr01 AS 'Peso Bruto', " &_
			"fra.p_adv102 AS 'P_Advalorem', " &_
			"IFNULL((SELECT SUM(adv.import36) " &_
			"FROM rku_extranet.sscont36 AS adv " &_
			"WHERE adv.refcia36 = ops.refcia01 AND adv.patent36 = ops.patent01 AND adv.adusec36 = ops.adusec01 AND adv.cveimp36 = 6 " &_
			"GROUP BY adv.refcia36), 0) AS 'Advalorem', " &_
			"IFNULL((SELECT SUM(iva.import36) " &_
			"FROM rku_extranet.sscont36 AS iva " &_
			"WHERE iva.refcia36 = ops.refcia01 AND iva.patent36 = ops.patent01 AND iva.adusec36 = ops.adusec01 AND iva.cveimp36 = 3 " &_
			"GROUP BY iva.refcia36), 0) AS 'IVA', " &_
			"IFNULL((SELECT SUM(dta.import36) " &_
			"FROM rku_extranet.sscont36 AS dta " &_
			"WHERE dta.refcia36 = ops.refcia01 AND dta.patent36 = ops.patent01 AND dta.adusec36 = ops.adusec01 AND dta.cveimp36 = 1 " &_
			"GROUP BY dta.refcia36), 0) AS 'DTA', " &_
			"IF(MID(ops.refcia01, 1, 3) IN ('TOL', 'DAI'), IFNULL((SELECT SUM(d21.mont21 * IF(e21.deha21 = 'C', -1, 1)) " &_
			"FROM rku_extranet.d21paghe AS d21 " &_
			"LEFT JOIN rku_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 " &_
			"LEFT JOIN rku_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 " &_
			"WHERE d21.refe21 = ops.refcia01 AND c21.desc21 LIKE '%MANIOBR%' " &_
			"GROUP BY e21.conc21), 0), 0) AS 'Maniobras', " &_
			"IFNULL((SELECT SUM(prv.import36) " &_
			"FROM rku_extranet.sscont36 AS prv " &_
			"WHERE prv.refcia36 = ops.refcia01 AND prv.patent36 = ops.patent01 AND prv.adusec36 = ops.adusec01 AND prv.cveimp36 = 15 " &_
			"GROUP BY prv.refcia36), 0) AS 'PRV', " &_
			"170 AS 'Validacion', " &_
			"IFNULL((SELECT SUM(e31.csce31) " &_
			"FROM rku_extranet.d31refer AS d31 " &_
			"INNER JOIN rku_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
			"WHERE d31.refe31 = ops.refcia01), 0) AS 'Serv_Complementarios', " &_
			"IFNULL((SELECT SUM(d21.mont21 * IF(e21.deha21 = 'C', -1, 1)) " &_
			"FROM rku_extranet.d21paghe AS d21 " &_
			"LEFT JOIN rku_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 " &_
			"LEFT JOIN rku_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 " &_
			"WHERE d21.refe21 = ops.refcia01 AND c21.desc21 LIKE '%DEMOR%' " &_
			"GROUP BY e21.conc21), 0) AS 'Demoras', " &_
			"IFNULL((SELECT SUM(eci.import36) " &_
			"FROM rku_extranet.sscont36 AS eci " &_
			"WHERE eci.refcia36 = ops.refcia01 AND eci.patent36 = ops.patent01 AND eci.adusec36 = ops.adusec01 AND eci.cveimp36 = 18 " &_
			"GROUP BY eci.refcia36), 0) AS 'ECI', " &_
			"0 AS 'Otros Pagos', " &_
			"IF(MID(ops.refcia01, 1, 3) NOT IN ('TOL', 'DAI'), IFNULL((SELECT SUM(d21.mont21 * IF(e21.deha21 = 'C', -1, 1)) " &_
			"FROM rku_extranet.d21paghe AS d21 " &_
			"LEFT JOIN rku_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 " &_
			"LEFT JOIN rku_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 " &_
			"WHERE d21.refe21 = ops.refcia01 AND c21.desc21 LIKE '%MANIOBR%' " &_
			"GROUP BY e21.conc21), 0), 0) AS 'Maniobras Puerto', " &_
			"IFNULL((SELECT SUM(e31.chon31) " &_
			"FROM rku_extranet.d31refer AS d31 " &_
			"INNER JOIN rku_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
			"WHERE d31.refe31 = ops.refcia01), 0) AS 'Honorarios', " &_
			"IFNULL((SELECT SUM(d21.mont21 * IF(e21.deha21 = 'C', -1, 1)) " &_
			"FROM rku_extranet.d21paghe AS d21 " &_
			"LEFT JOIN rku_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 " &_
			"LEFT JOIN rku_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 " &_
			"WHERE d21.refe21 = ops.refcia01 AND c21.desc21 LIKE '%RECO%ADUA%' " &_
			"GROUP BY e21.conc21), 0) AS 'Reconocimiento Aduanero', " &_
			"IFNULL((SELECT SUM(e31.chon31 * 0.16) " &_
			"FROM rku_extranet.d31refer AS d31 " &_
			"INNER JOIN rku_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
			"WHERE d31.refe31 = ops.refcia01), 0) AS 'IVA Honorarios', " &_
			"IFNULL((SELECT GROUP_CONCAT(DISTINCT(e31.cgas31)) " &_
			"FROM rku_extranet.d31refer AS d31 " &_
			"INNER JOIN rku_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
			"WHERE d31.refe31 = ops.refcia01), 0) AS 'Cuenta de Gastos', " &_
			"IFNULL((SELECT SUM(e31.tota31) " &_
			"FROM rku_extranet.d31refer AS d31 " &_
			"INNER JOIN rku_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
			"WHERE d31.refe31 = ops.refcia01), 0) AS 'Importe Cuenta de Gastos', " &_
			"IFNULL((SELECT GROUP_CONCAT(c20.nomb20) " &_
			"FROM rku_extranet.d21paghe AS d21 " &_
			"LEFT JOIN rku_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 " &_
			"LEFT JOIN rku_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 " &_
			"LEFT JOIN rku_extranet.c20benef AS c20 ON c20.clav20 = e21.bene21 AND c20.aplic20 = 'T' " &_
			"WHERE d21.refe21 = ops.refcia01 AND c21.desc21 LIKE '%FLETE%' " &_
			"GROUP BY e21.conc21), '') AS 'transportista', " &_
			"'' AS 'Flete Terrestre', " &_
			"'' AS 'Naviera', " &_
			"'' AS 'Flete Maritimo', " &_
			"'' AS 'Linea Aerea', " &_
			"'' AS 'Flete aereo', " &_
			"fra.paiscv02 AS 'Pais Vendedor', " &_
			"fra.paiori02 AS 'Pais Origen', " &_
			"IFNULL((SELECT DISTINCT IF(par.cveide12 IS NOT NULL AND par.cveide12 <> '', 'Si', 'No') " &_
			"FROM rku_extranet.ssipar12 AS par " &_
			"WHERE par.refcia12 = ops.refcia01 AND par.patent12 = ops.patent01 AND par.adusec12 = ops.adusec01 AND par.ordfra12 = fra.ordfra02 " &_
			"AND par.cveide12 IN ('AL', 'CM', 'CP', 'EF', 'PS', 'TL')), 'No') AS 'Beneficio Arancelario', " &_
			"(SELECT GROUP_CONCAT(DISTINCT(fac.terfac39)) " &_
			"FROM rku_extranet.ssfact39 AS fac " &_
			"WHERE fac.refcia39 = ops.refcia01 AND fac.patent39 = ops.patent01 AND fac.adusec39 = ops.adusec01 " &_
			"GROUP BY fac.refcia39) AS 'Incoterm', " &_
			"fra.vincul02 AS 'Vinculacion', " &_
			"fra.metval02 AS 'Valoracion', " &_
			"ops.otros01 AS 'Incrementables', " &_
			"ops.decble01 AS 'Decrementables', " &_
			"IFNULL((SELECT DISTINCT IF(rc.refcia06 IS NOT NULL AND rc.refcia06 <> '', 'Si', 'No') " &_
			"FROM rku_extranet.ssrecp06 AS rc " &_
			"WHERE rc.reforg06 = ops.refcia01), 'No') AS 'Rectificacion', " &_
			"'' AS 'Fecha', " &_
			"'' AS 'NUM_POLIZA', " &_
			"'' AS 'TIPO_POLIZA', " &_
			"'' AS 'IMPORTE_MXP', " &_
			"'' AS 'CARPETA', " &_
			"IFNULL((SELECT GROUP_CONCAT(DISTINCT(con.numcon40) SEPARATOR ', ') " &_
			"FROM rku_extranet.sscont40 AS con " &_
			"WHERE con.refcia40 = ops.refcia01 AND con.patent40 = ops.patent01 AND con.adusec40 = ops.adusec01 " &_
			"GROUP BY con.refcia40), '') AS 'Contenedores' " &_
			"FROM rku_extranet.ssdagi01 AS ops " &_
			"INNER JOIN rku_extranet.ssconf32 AS con ON con.patent32 = ops.patent01 AND con.cveadu32 = ops.cveadu01 AND con.cvesec32 = ops.cvesec01 " &_
			"LEFT JOIN rku_extranet.ssrecp06 AS rec ON rec.refcia06 = ops.refcia01 AND rec.patent06 = ops.patent01 " &_
			"AND rec.adusec06 = ops.adusec01 " &_
			"LEFT JOIN rku_extranet.ssfrac02 AS fra ON fra.refcia02 = ops.refcia01 AND fra.patent02 = ops.patent01 AND fra.adusec02 = ops.adusec01 " &_
			"WHERE ops.firmae01 <> '' AND ops.firmae01 IS NOT NULL " &_
			"AND ops.fecpag01 >= '" & DateI & "' AND ops.fecpag01 <= '" & DateF & "' " &_
			"AND ops.rfccli01 = 'CME8909276S1' " &_
			"ORDER BY ops.refcia01, fra.ordfra02; "
	 ' Response.Write(SQL)
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
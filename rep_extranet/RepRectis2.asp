<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=1500


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

tOpe = Request.Form("TipoOp")

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
		case "TOL"
			strOficina="tol"
	end select

	cve=request.form("cve")
	mov=request.form("mov")
	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	Vrfc=Request.Form("rfcCliente")
	'Response.Write(Vrfc)
	Vckcve=Request.Form("ckcve")
	tOpe = Request.Form("TipoOp")
	Vclave=Request.Form("txtCliente")

	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	MesIn = UCase(MonthName(Month(fi)))
	DateI = "'" & Anioi & "/" & Mesi & "/" & Diai & "'"

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	MesFi = UCase(MonthName(Month(ff)))
	DateF = "'" & AnioF & "/" & MesF & "/" & DiaF & "'"
	
	nocolumns = 9
	query = GeneraSQL
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write("DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	 'Response.Write(query)
	 'Response.End()
	Set RSops = CreateObject("ADODB.RecordSet")
	'response.write(query)
	'response.end()
	Set RSops = ConnStr.Execute(query)
	'Response.Write(RSops)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
	
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""850""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"REPORTE DE RECTIFICACIONES" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>" &_
												"IMPORTACI&Oacute;N / EXPORTACI&Oacute;N" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center> " &_
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>"
												if AnioI = AnioF then
													if MesIn = MesFi then
														info = info & "DEL " & DiaI & " AL " & DiaF & " DE " & MesFi & " DE " & AnioF
													else
														info = info & "DEL " & DiaI & " DE " & MesIn & " AL " & DiaF & " DE " & MesFi & " DEL " & AnioF
													end if
												else
													info = info & "DEL " & DiaI & " DE " & MesIn & " DE " & AnioI & " AL " & DiaF & " DE " & MesFi & " DE " & AnioF
												end if
											info = info & "</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
								"</td>" &_
							"</tr>" &_
				"</table><br>"
		
		header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">" &_
								celdahead("Operaci&oacute;n") &_
								celdahead("Oficina") &_
								celdahead("Pedimento de Rectificaci&oacute;n") &_
								celdahead("Referencia de Rectificaci&oacute;n") &_
								celdahead("Pedimento Original") &_
								celdahead("Referencia Original") &_
								celdahead("Fecha de Pago de Rectificaci&oacute;n") &_
								celdahead("Nombre del Cliente") &_
								celdahead("RFC de Cliente")
								
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
								celdadatos(RSops.Fields.Item("Operacion").Value) &_
								celdadatos(RSops.Fields.Item("Oficina").Value) &_
								celdadatos(RSops.Fields.Item("PedRec").Value) &_
								celdadatos(RSops.Fields.Item("RefRec").Value) &_
								celdadatos(RSops.Fields.Item("Original").Value) &_
								celdadatos(RSops.Fields.Item("RefOri").Value) &_
								celdadatos(RSops.Fields.Item("FRecti").Value) &_
								celdadatos(RSops.Fields.Item("Cliente").Value) &_
								celdadatos(RSops.Fields.Item("RFC").Value)		
			datos = datos &	"</tr>"
			Rsops.MoveNext()
			
		Loop
	'Response.Write(datos)	
	'Response.Write(RSops.Fields.Item("Operacion").Value)
	'prom = ""
	'prom = Promedios
	'Response.Write(info & header & datos & "</table><br>")
	' Response.End()
	html = info & header & datos & "</table><br>" 
	'& prom & "<br>" & Actualizaciones
	
	
	End If
end if

Function Promedios
	SQLpromedios = ""
	condicion = filtro
	SQLpromedios = 						"SELECT  i.refcia01 AS 'referencia', "
	if mov = "i" Then
		SQLpromedios = SQLpromedios & 	kpi("AVG", "i.fecent01", "c.fdsp01") & "as 'AVGCTE', " &_
										kpi("MAX", "i.fecent01", "c.fdsp01") & "as 'MAXCTE', " &_
										kpi("MIN", "i.fecent01", "c.fdsp01") & "as 'MINCTE', "
	Else
		SQLpromedios = SQLpromedios &	kpi("AVG", "i.fecpre01", "c.fdsp01") & "as 'AVGCTE', " &_
										kpi("MAX", "i.fecpre01", "c.fdsp01") & "as 'MAXCTE', " &_
										kpi("MIN", "i.fecpre01", "c.fdsp01") & "as 'MINCTE', "
	End If	
	SQLpromedios = SQLpromedios & 		kpi("AVG", "c.frev01", "c.fdsp01") & "as 'AVGGRK', " &_
										kpi("MAX", "c.frev01", "c.fdsp01") & "as 'MAXGRK', " &_
										kpi("MIN", "c.frev01", "c.fdsp01") & "as 'MINGRK', " &_
										kpi("AVG", "c.fdsp01", "cta.fech31") & "as 'AVGADMIN', " &_
										kpi("MAX", "c.fdsp01", "cta.fech31") & "as 'MAXADMIN', " &_
										kpi("MIN", "c.fdsp01", "cta.fech31") & "as 'MINADMIN', " &_
										kpi("AVG", "cta.fech31", "cta.frec31") & "as 'AVGACUSE', " &_
										kpi("MAX", "cta.fech31", "cta.frec31") & "as 'MAXACUSE', " &_
										kpi("MIN", "cta.fech31", "cta.frec31") & "as 'MINACUSE' " &_
										"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
										"LEFT JOIN " & strOficina & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 " &_
										"LEFT JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
										"LEFT JOIN " & strOficina & "_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " &_
										"LEFT JOIN " & strOficina & "_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31 " &_
										"WHERE i.firmae01 IS NOT NULL AND i.firmae01 <> '' AND i.cveped01 <> 'R1' " &_
										"AND c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "' " & condicion &_
										"AND (cta.esta31 <> 'C' ) " &_
										"AND (cta.fech31 >= c.fdsp01 Or cta.fech31 IS NOT NULL) " &_
										"GROUP BY MID(i.refcia01,1,3) " &_
										"ORDER BY i.refcia01"
	' Response.Write(SQLpromedios)
	' Response.End
	Set RSprom = CreateObject("ADODB.RecordSet")
	Set RSprom = ConnStr.Execute(SQLpromedios)
	'Response.write(SQLpromedios)
	'Response.End()
	RSprom.MoveFirst()
	construc = ""
	construc = 					"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
									"<tr bgcolor = ""#006699"" class = ""boton"">" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												"<td>" &_
												"</td>" &_
												celdahead("Promedio") &_
												celdahead("Maximo") &_
												celdahead("Minimo") &_
											"</font>" &_
										"</strong>" &_
									"</tr>" &_
									"<tr>" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												celdahead("Despacho - Entrada") &_
												celdadatos(RSprom.Fields.Item("AVGCTE").Value) &_
												celdadatos(RSprom.Fields.Item("MAXCTE").Value) &_
												celdadatos(RSprom.Fields.Item("MINCTE").Value) &_
											"</font>" &_
										"</strong>" &_
									"</tr>"
	if mov = "i" then
		construc = construc & 		"<tr>" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												celdahead("Despacho - Revalidacion") &_
												celdadatos(RSprom.Fields.Item("AVGGRK").Value) &_
												celdadatos(RSprom.Fields.Item("MAXGRK").Value) &_
												celdadatos(RSprom.Fields.Item("MINGRK").Value) &_
											"</font>" &_
										"</strong>" &_
									"</tr>"
	End If
	construc = construc & 			"<tr>" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												celdahead("CG - Despacho") &_
												celdadatos(RSprom.Fields.Item("AVGADMIN").Value) &_
												celdadatos(RSprom.Fields.Item("MAXADMIN").Value) &_
												celdadatos(RSprom.Fields.Item("MINADMIN").Value) &_
											"</font>" &_
										"</strong>" &_
									"</tr>" &_
									"<tr>" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												celdahead("Acuse - CG") &_
												celdadatos(RSprom.Fields.Item("AVGACUSE").Value) &_
												celdadatos(RSprom.Fields.Item("MAXACUSE").Value) &_
												celdadatos(RSprom.Fields.Item("MINACUSE").Value) &_
											"</font>" &_
										"</strong>" &_
									"</tr>" &_
								"</table>"
	'Response.Write(construc)
	'Response.End()
	RSprom.Close()
	Set RSprom = Nothing
	Promedios = construc
End Function

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
	If IsNull(texto) = True Or texto = "" Then
		texto = "No Capturado"
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
'function filtro
	'verifica si esta seleccionado la opcion rfc y le asigna el valor del rfc a condicion para agregarlo a la cadena SQL
'	if Vckcve = 0 then
'		if IsNumeric(Vrfc) then
'			if Vrfc=0 then
'				condicion=""
'			End if
'		else
'		condicion = "AND i.rfccli01 = '" & Vrfc & "' "
'		end if
'	end if
	
'	filtro = condicion
'end function

function GeneraSQL
	condicion = filtro
	tabla=""
	Ope=""
	'compara tipo de operacion y asigna la tabla y el nombre de la operacion
	if tOpe=1 then
		tabla="ssdagi01"
		Ope="'Impo'"
	else
		tabla="ssdage01"
		Ope="'Expo'"
	end if
		SQL = "select " & Ope & " as Operacion , mid(i.refcia01 ,1,3) as Oficina,CONCAT (i.numped01, '-' ,i.adusec01,'-',i.patent01 ) as PedRec ," &_
			"rec.refcia06 as RefRec,rec.pedorg06 as Original, rec.reforg06 as RefOri,rec.f_pagr06 as FRecti, i.nomcli01 as Cliente, i.rfccli01 as RFC " &_
			"from "& strOficina &"_extranet."& tabla &" as i " &_
			"inner join "& strOficina &"_extranet.ssrecp06 as rec on rec.refcia06 =i.refcia01 " &_
			"where i.fecpag01 >=" & DateI & " and i.fecpag01 <=" & DateF & " and i.firmae01 <>'' " & condicion &_
			"group by rec.refcia06 Order by i.nomcli01"
		GeneraSQL = SQL
		'response.write(SQL)
		'response.end()
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


Function Actualizaciones()
	html = ""
	cont = 0
	log_act =	"SELECT 'RKU' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM rku_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'DAI' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM dai_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'SAP' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM sap_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'LZR' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM lzr_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'CEG' as Ofi, max(d_fechahora_act) as fecha " &_
				"FROM ceg_extranet.log_actualiza " &_
				"group by ofi " &_
				"UNION ALL " &_
				"SELECT 'TOL' as Ofi, max(d_fechahora_act) as fecha " &_
				"FROM tol_extranet.log_actualiza " &_
				"group by ofi " &_
				"order by ofi "
	
	Set RSact = CreateObject("ADODB.RecordSet")
	Set RSact = ConnStr.Execute(log_act)
	RsAct.MoveFirst
	
	
	html = html &	"<table border='2' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
						"<tr bgcolor = ""#006699"" class = ""boton"">" &_
							"<td colspan=4>" &_
								"<center>" &_
									"<strong>" &_
										"<font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
											"Ultimas Actualizaciones" &_
										"</font>" &_
									"</strong>" &_
								"</center>" &_
							"</td>" &_
						"</tr>" &_
						"<tr>"
		
	 Do Until RsAct.EOF = true
		html = html & 		"<td>" & RsAct("ofi") & "</td>" &_
							"<td>" & RsAct("fecha") & "</td>"
		cont = cont + 1
		if cont = 2 then
			html = html & "</tr><tr>"
			cont = 0
		End If
		RsAct.MoveNext
	Loop
	
	html = html & 		"</tr>" &_
					"</table><br><br>"
	RSAct.Close()
	Set RSAct = Nothing
	Actualizaciones = html
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
		<TITLE>::.... REPORTE DE RECTIFICACIONES .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
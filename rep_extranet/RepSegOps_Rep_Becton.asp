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
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	blnAplicaFiltro = true
end if
if blnAplicaFiltro then
	permi = " AND cvecli01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	permi = ""
end if


Tiporepo = Request.Form("TipRep")
if Tiporepo = 2 Then
	Response.Addheader "Content-Disposition", "attachment;"
	Response.ContentType = "application/vnd.ms-excel"
End If


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
	nocols = 0
	tablamov = ""
	if mov = "i" then
		movi = ":: IMPORTACION ::"
		tablamov = "ssdagi01"
		nocols = 28
		query = GeneraSQL
	else
		movi = ":: EXPORTACION ::"
		tablamov = "ssdage01"
		nocols = 22
		query = GeneraSQL
	end if
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	' Response.Write(query)
	' Response.End()
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocols & """>" &_
											"<p align=""left"">" &_
												"::.... REPORTE DE SEGUIMIENTO DE OPERACIONES BD .... ::" &_
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
		
		header = 			"<tr bgcolor = ""#006699"" class = ""boton"">" &_
								celdahead("Ind") &_
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Cliente") &_
								celdahead("Ejecutivo")
		if mov = "i" Then
			header = header & 	celdahead("Descrip. Mercancia") &_
								celdahead("Pais Origen") &_
								celdahead("Contenedores")
		Else
			header = header & 	celdahead("No. Factura") &_
								celdahead("Pais Destino")
		End If
		header = header & 		celdahead("Bultos")
		if mov = "i" Then
			header = header &	celdahead("ETA") &_
								celdahead("Documentos") &_
								celdahead("Revalidacion") &_
								celdahead("Previo") &_
								celdahead("PagPedto") &_	
								celdahead("Despacho") &_
								celdahead("Entrada") &_
								celdahead("IndDsp") &_
								celdahead("IndDsp2")
		else
			header = header &	celdahead("Documentos") &_
								celdahead("Fecha de Entrada") &_
								celdahead("PagPedto") &_
								celdahead("Despacho") &_
								celdahead("IndDsp")
		end if
		header = header &		celdahead("Observaciones Trafico")
		if mov = "i" Then
			header = header & 	celdahead("VacioCont")
		end if
		header = header & 		celdahead("Fecha C.Gastos") &_
								celdahead("C.Gastos") &_
								celdahead("IndCGast") &_
								celdahead("Fec.Acuse Recibo") &_
								celdahead("IndAcuse") &_
								celdahead("Observaciones Administraci&oacute;n") &_
								celdahead("Semaforo") &_
								celdahead("Muestreo") &_
							"</tr>"
		datos = ""
		contador = 1
		facturas = ""
		referencia = ""
		fechalib = ""
		Do Until RSops.EOF
			referencia = RSops.Fields.Item("referencia")
			if mov = "i" Then
				contenedores = contienecont(referencia)
				fechalib = liberacont(referencia)
			Else
				facturas = contienefacturas(referencia)
			End If
			'contenedores = contienecont("RKU10-02622")
			datos = datos &	"<tr>" &_
								celdadatos(contador) &_
								celdadatos(referencia) &_
								celdadatos(RSops.Fields.Item("pedimento").Value) &_
								celdadatos(RSops.Fields.Item("cvecli").Value) &_
								celdadatos(RSops.Fields.Item("ejecutivo").Value)
			if mov = "i" Then
				datos = datos &	celdadatos(RSops.Fields.Item("Descpro").Value)
			Else
				datos = datos &	celdadatos(facturas)
			End If
			datos = datos &		celdadatos(RSops.Fields.Item("porigen").Value)
			if mov = "i" Then
				datos = datos & celdadatos(contenedores)
			End If
			datos = datos &		celdadatos(RSops.Fields.Item("bultos").Value)
			if mov = "i" Then
				datos = datos &	celdadatos(RSops.Fields.Item("feta").Value) &_
								celdadatos(RSops.Fields.Item("fdocs").Value) &_
								celdadatos(RSops.Fields.Item("frev").Value) &_
								celdadatos(RSops.Fields.Item("fprev").Value) &_
								celdadatos(RSops.Fields.Item("fpago").Value) &_
								celdadatos(RSops.Fields.Item("fdesp").Value) &_
								celdadatos(RSops.Fields.Item("fentrada").Value) &_
								celdadatos(RSops.Fields.Item("KPIGRK").Value)
			else
				datos = datos &	celdadatos(RSops.Fields.Item("fdocs").Value) &_
								celdadatos(RSops.Fields.Item("fentrada").Value) &_
								celdadatos(RSops.Fields.Item("fpago").Value) &_
								celdadatos(RSops.Fields.Item("fdesp").Value)
			end if
			datos = datos & 	celdadatos(RSops.Fields.Item("KPICTE").Value) &_
								celdadatos(Observaciones(referencia))
			if mov = "i" Then
				datos = datos &	celdadatos(fechalib)
			End If
			datos = datos &		celdadatos(RSops.Fields.Item("FCG").Value) &_
								celdadatos(RSops.Fields.Item("CG").Value) &_
								celdadatos(RSops.Fields.Item("KPIADMIN").Value) &_
								celdadatos(RSops.Fields.Item("facuse").Value) &_
								celdadatos(RSops.Fields.Item("KPIACUSE").Value) &_
								celdadatos("&nbsp;") &_
								celdadatos(RSops.Fields.Item("semaforo").Value) &_
								celdadatos(RSops.Fields.Item("muestra").Value) &_
							"</tr>"
							
			contador = contador + 1
			Rsops.MoveNext()
		Loop
	
	prom = ""
	prom = Promedios
	' Response.Write(info & header & datos & "</table><br>" & prom)
	' Response.End()
	html = Actualizaciones & info & header & datos & "</table><br>" & prom
	
	
	End If
end if

Function Promedios
	SQLpromedios = ""
	condicion = filtro
	SQLpromedios = 						"SELECT  i.refcia01 AS 'referencia', "
	if mov = "i" Then
		SQLpromedios = SQLpromedios & 	KPI("AVG", "i.fecent01", "c.fdsp01") & " as 'AVGCTE', " &_
										KPI("MAX", "i.fecent01", "c.fdsp01") & " as 'MAXCTE', " &_
										KPI("MIN", "i.fecent01", "c.fdsp01") & " as 'MINCTE', "
	Else
		SQLpromedios = SQLpromedios &	KPI("AVG", "c.frec01", "c.fdsp01") & " as 'AVGCTE', " &_
										KPI("MAX", "c.frec01", "c.fdsp01") & " as 'MAXCTE', " &_
										KPI("MIN", "c.frec01", "c.fdsp01") & " as 'MINCTE', "
	End If	
	SQLpromedios = SQLpromedios & 		KPI("AVG", "c.frev01", "c.fdsp01") & " as 'AVGGRK', " &_
										KPI("MAX", "c.frev01", "c.fdsp01") & " as 'MAXGRK', " &_
										KPI("MIN", "c.frev01", "c.fdsp01") & " as 'MINGRK', " &_
										KPI("AVG", "c.fdsp01", "cta.fech31") & " as 'AVGADMIN', " &_
										KPI("MAX", "c.fdsp01", "cta.fech31") & " as 'MAXADMIN', " &_
										KPI("MIN", "c.fdsp01", "cta.fech31") & " as 'MINADMIN', " &_
										KPI("AVG", "cta.fech31", "cta.frec31") & " as 'AVGACUSE', " &_
										KPI("MAX", "cta.fech31", "cta.frec31") & " as 'MAXACUSE', " &_
										KPI("MIN", "cta.fech31", "cta.frec31") & " as 'MINACUSE' " &_
										"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
										"LEFT JOIN " & strOficina & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 " &_
										"LEFT JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
										"LEFT JOIN " & strOficina & "_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " &_
										"LEFT JOIN " & strOficina & "_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31 " &_
										"LEFT JOIN trackingbahia.bit_soia as bs on i.refcia01 = bs.frmsaai01 and (bs.detsit01 = 510 or bs.detsit01 = 310) " &_
										"WHERE i.firmae01 IS NOT NULL AND i.firmae01 <> '' AND i.cveped01 <> 'R1' " &_
										"AND c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "' " & condicion &_
										"AND (cta.esta31 <> 'C' or cta.esta31 IS NULL) " &_
										"AND (cta.fech31 >= c.fdsp01 Or cta.fech31 IS NULL) " &_
										"GROUP BY MID(i.refcia01,1,3) " &_
										"ORDER BY i.cvecli01, i.refcia01"
	' Response.Write(SQLpromedios)
	' Response.End
	Set RSprom = CreateObject("ADODB.RecordSet")
	Set RSprom = ConnStr.Execute(SQLpromedios)
	RSprom.MoveFirst()
	construc = ""
	construc = 					"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
									"<tr bgcolor = ""#006699"" class = ""boton"">" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												"<td>" &_
													"&nbsp;" &_
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
									"</tr>" &_
									"<tr>" &_
										"<strong>" &_
											"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">"
	if mov = "i" then
		construc = construc & 					celdahead("Despacho - Revalidacion") &_
												celdadatos(RSprom.Fields.Item("AVGGRK").Value) &_
												celdadatos(RSprom.Fields.Item("MAXGRK").Value) &_
												celdadatos(RSprom.Fields.Item("MINGRK").Value)
	' Else
		' construc = construc &					celdadatos(RSprom.Fields.Item("AVGCTE").Value) &_
												' celdadatos(RSprom.Fields.Item("MAXCTE").Value) &_
												' celdadatos(RSprom.Fields.Item("MINCTE").Value)
	End If
	construc = construc & 					"</font>" &_
										"</strong>" &_
									"</tr>" &_
									"<tr>" &_
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
	Set RSprom = Nothing
	Promedios = construc
End Function

function celdahead(texto)
	cell = "<td bgcolor = ""#006699"" width=""100"" nowrap>" &_
				"<strong>" &_
					"<font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
						texto &_
					"</font>" &_
				"</strong>" &_
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
		condicion = "AND i.cvecli01 = " & Vclave & " "
	end if
	filtro = condicion
end function

function GeneraSQL
	SQL = ""
	condicion = filtro
	SQL = 	"SELECT  i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pedimento', " &_
			"i.cvecli01 AS 'cvecli', " &_
			"i.nomcli01 AS 'nomcli', " &_
			"fr.d_mer102 AS 'Descpro', " &_
			"fr.paiori02 AS 'porigen', " &_
			"i.totbul01 AS 'bultos', " &_
			"IF(c.feta01 IS NULL OR c.feta01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.feta01,'%d-%m-%Y')) AS 'feta',  " &_
			"IF(c.fdoc01 IS NULL OR c.fdoc01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.fdoc01,'%d-%m-%Y')) AS 'fdocs',  " &_
			"IF(c.frev01 IS NULL OR c.frev01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.frev01,'%d-%m-%Y')) AS 'frev',  " &_
			"IF(c.fpre01 IS NULL OR c.fpre01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.fpre01,'%d-%m-%Y')) AS 'fprev',  " &_
			"IF(c.fdsp01 IS NULL OR c.fdsp01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.fdsp01,'%d-%m-%Y')) AS 'fdesp',  " &_
			"IF(i.fecpag01 IS NULL OR i.fecpag01 = '0000-00-00', 'No Capturada', DATE_FORMAT(i.fecpag01,'%d-%m-%Y')) AS 'fpago', "
			if mov = "i" Then
				SQL = SQL & "IF(i.fecent01 IS NULL OR i.fecent01 = '0000-00-00', 'No Capturada', DATE_FORMAT(i.fecent01,'%d-%m-%Y')) AS 'fentrada', " &_
				KPI("", "i.fecent01", "c.fdsp01") & " as 'KPICTE', "
			Else
				SQL = SQL & "IF(c.frec01 IS NULL OR c.frec01 = '0000-00-00' , 'No existe en eZego', DATE_FORMAT(c.frec01, '%d-%m-%Y')) as 'fentrada', " &_
				KPI("", "c.frec01", "c.fdsp01") & " as 'KPICTE', "
			end if
	SQL = SQL & KPI("","c.frev01","c.fdsp01") & " as 'KPIGRK',  " &_
			"IF(cta.fech31 IS NULL OR c.fdsp01 = '0000-00-00','No Hay CG',DATE_FORMAT(cta.fech31,'%d-%m-%Y')) as 'FCG',  " &_
			"IF(cta.cgas31 IS NULL OR cta.cgas31 = '', 'No se ha Facturado',cta.cgas31) as 'CG',  " &_
			KPI("", "c.fdsp01", "cta.fech31") & " as 'KPIADMIN', " &_
			"IF(cta.frec31 IS NULL OR cta.frec31 = '0000-00-00', 'No Capturada', DATE_FORMAT(cta.frec31, '%d-%m-%Y')) as 'facuse', " &_
			KPI("", "cta.fech31", "cta.frec31") & " as 'KPIACUSE', " &_
			"IF(bs.detsit01 IS NULL or bs.detsit01 = '', 'Verde', 'Rojo') as 'semaforo', " &_
			"'' as 'muestra', " &_
			"d18.ejec18 AS 'ejecutivo' " &_
			"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
			"LEFT JOIN " & strOficina & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 " &_
			"LEFT JOIN " & strOficina & "_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 and d18.clie18 = c.clie01 " &_
			"LEFT JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
			"LEFT JOIN " & strOficina & "_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " &_
			"LEFT JOIN " & strOficina & "_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31 " &_
			"LEFT JOIN trackingbahia.bit_soia as bs on i.refcia01 = bs.frmsaai01 and (bs.detsit01 = 510 or bs.detsit01 = 310) " &_
			"WHERE i.firmae01 IS NOT NULL AND i.firmae01 <> '' AND i.cveped01 <> 'R1' " &_
			"AND ((c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "') OR c.fdsp01 is null)" & condicion &_
			"AND (cta.esta31 <> 'C' or cta.esta31 IS NULL) " &_
			"AND (cta.fech31 >= c.fdsp01 Or cta.fech31 IS NULL) " &_
			"ORDER BY i.cvecli01, i.refcia01"
			 ' Response.Write(SQL)
			 ' Response.End
	GeneraSQL = SQL
end function


Function contienecont(refe)
	sqlcont = "SELECT numcon40 FROM " & strOficina & "_extranet.sscont40 WHERE refcia40 = '" & refe & "'"
	conte = ""
	Set RScont = CreateObject("ADODB.RecordSet")
	Set RScont = ConnStr.Execute(sqlcont)
	IF RScont.EOF = True And RScont.BOF = True Then
		conte = "&nbsp;"
	Else
		RScont.MoveFirst
		Do Until RScont.EOF
			conte = conte & RScont.Fields.Item("numcon40").Value & ", "
			RScont.MoveNext
		Loop
		conte = MID(conte,1,LEN(conte)-2)
	End If
	RScont.Close()
	Set RScont = Nothing
	contienecont = conte
end function

Function liberacont(refe)
	sqlcont = 	"SELECT DATE_FORMAT(MAX(libc01), '%d-%m-%Y') as fecha FROM " &_
				strOficina & "_extranet.d01conte " &_
				"WHERE refe01 = '" & refe & "' GROUP BY refe01"
	' Response.Write(sqlcont)
	' Response.End()
	conte = ""
	Set RScont = CreateObject("ADODB.RecordSet")
	Set RScont = ConnStr.Execute(sqlcont)
	IF RScont.EOF = True And RScont.BOF = True Then
		conte = "&nbsp;"
	Else
		RScont.MoveFirst
		Do Until RScont.EOF
			conte = RScont.Fields.Item("fecha").Value
			RScont.MoveNext
		Loop
	End If
	RScont.Close()
	Set RScont = Nothing
	liberacont = conte
end function


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
				"GROUP BY ofi " &_
				"ORDER BY ofi "
	
	Set RSact = CreateObject("ADODB.RecordSet")
	Set RSact = ConnStr.Execute(log_act)
	RsAct.MoveFirst
	
	
	html = html &	"<table border='1' align='center' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
						"<tr>" &_
							"<td colspan=4><center>Ultimas Actualizaciones</center></td>" &_
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
	
	RsAct.Close()
	Set RsAct = Nothing
	
	Actualizaciones = html
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

%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE SEGUIMIENTO DE OPERACIONES BD .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
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
' Response.Write(permi)
' Response.End



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


cve=request.form("cve")
mov=request.form("mov")
fi=trim(request.form("fi"))
ff=trim(request.form("ff"))
DiaI = cstr(datepart("d",fi))
Mesi = cstr(datepart("m",fi))
AnioI = cstr(datepart("yyyy",fi))
DateI = Anioi & "/" & Mesi & "/" & Diai
DiaF = cstr(datepart("d",ff))
MesF = cstr(datepart("m",ff))
AnioF = cstr(datepart("yyyy",ff))
DateF = AnioF & "/" & MesF & "/" & DiaF
Vrfc = Request.Form("rfcCliente")
Vckcve = Request.Form("ckcve")
Vclave = Request.Form("txtCliente")
nivel = request.Form("nivel")
etap = request.Form("etapas")
' Response.Write(etap)
' Response.End
if etap <> "T" Then
	filtroetapa = "AND ((m_observ IS NOT NULL AND m_observ <> '') OR etx.clavec <> '') "
End If

If Vckcve = 0 Then 
	filtrocliente = "AND i.rfccli01 = '" & Vrfc & "' "
Else 
	If Vclave = "Todos" Then 
		filtrocliente = permi
	Else
		filtrocliente = "AND i.cvecli01 = " & Vclave & " "
	End If
End If
' response.write("filtro cliente = " & filtrocliente)
' response.end()

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
	
	query = ""
	tablamov = ""
	if mov = "i" then
		movi = ":: IMPORTACION ::"
		tablamov = "ssdagi01"
		query = GeneraSQL
	else
		movi = ":: EXPORTACION ::"
		tablamov = "ssdage01"
		query = GeneraSQL
	end if
	Set ConnStr = Server.CreateObject("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = Server.CreateObject("ADODB.Recordset")
	Set RSops = ConnStr.Execute(query)
	if RSops.Eof = True and RSops.Bof = True Then
		Response.Write(	"<table border=""0"" align = ""center"" cellpadding = ""0"" cellspacing = ""7"" class = ""titulo1"">" &_
							"<tr>" &_
								"<td>" &_
									"No Existen Datos Que Cumplan Esos Requerimientos" &_
								"</td>" &_
							"</tr>" &_
						"</table>")
		Response.End()
	Else
		If Tiporepo = "excel" Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		RSops.MoveFirst
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
						"<tr>" &_
							"<strong>" &_
								"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
									"<td colspan=""22"">" &_
										"<p align=""left"">" &_
											"::.... REPORTE DE SEGUIMIENTO DE OPERACIONES .... ::" &_
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
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Clave Cliente") &_
								celdahead("Cliente") &_
								celdahead("Ejecutivo") &_
								celdahead("Descrip. Mercancia") &_
								celdahead("Contenedores") &_
								celdahead("Bultos") &_
								celdahead("ETA") &_
								celdahead("Documentos") &_
								celdahead("Revalidacion") &_
								celdahead("Previo") &_
								celdahead("PagPedto") &_	
								celdahead("Despacho")
		if mov = "i" Then
			header = header &	celdahead("Entrada")
		Else
			header = header & 	celdahead("Alta Ref")
		End If
		header = header & 		celdahead("KPI Despacho - Entrada") &_
								celdahead("KPI Despacho - Revalidacion") &_
								celdahead("Fecha C.Gastos") &_
								celdahead("C.Gastos") &_
								celdahead("KPI Fecha GC - Despacho") &_
								celdahead("Fec.Acuse Recibo") &_
								celdahead("KPI Fecha Acuse Recibo - Fecha CG") &_
							"</tr>"
		' Response.Write(info & header)
		' Response.End
		datos = ""
		referencia = ""
		etapa = ""
		depu=""
		cont = 0
		Set RSmerc = Server.CreateObject("ADODB.Recordset")
		Set RScont = Server.CreateObject("ADODB.Recordset")
		Do Until RSops.Eof
			datos = datos & "<tr>"
			If referencia <> RSops("referencia") Then
				referencia = RSops("referencia")
				datos = datos &	celdadatos(referencia) &_
								celdadatos(RSops("pedimento")) &_
								celdadatos(RSops("cvecli")) &_
								celdadatos(RSops("nomcli")) &_
								celdadatos("&nbsp;")
				fracciones = ""
				query = "SELECT fr.refcia02, fr.fraarn02, fr.d_mer102 FROM " & strOficina & "_extranet.ssfrac02 AS fr WHERE refcia02 = '" & referencia & "'"
				Set RSmerc = ConnStr.Execute(query)
				If RSmerc.Eof = True and RSmerc.Bof = True Then
					fracciones = "No capturado"
				Else
					Do Until RSmerc.eof
					fracciones = RSmerc("fraarn02") & " " & RSmerc("d_mer102") & ", "
					RSmerc.MoveNext()
					Loop
					fracciones = MID(fracciones,1,LEN(fracciones)-2)
				End if
				query = "SELECT con.refcia40, con.numcon40 FROM " & strOficina & "_extranet.sscont40 AS con WHERE refcia40 = '" & referencia & "'"
				' response.write(query)
				' response.end()
				conte = ""
				Set RScont = ConnStr.Execute(query)
				If RScont.Eof = True and RScont.Bof = True Then
					conte = "No capturado"
				Else
					Do Until RScont.Eof
					conte = RScont("numcon40") & ", "
					RScont.MoveNext
					Loop
					conte = MID(conte,1,LEN(conte)-2)
				End if
				
				datos = datos &	celdadatos(fracciones) &_
								celdadatos(conte) &_
								celdadatos(RSops("bultos")) &_
								celdadatos(RSops("feta")) &_
								celdadatos(RSops("fdocs")) &_
								celdadatos(RSops("frev")) &_
								celdadatos(RSops("fprev")) &_
								celdadatos(RSops("fpago")) &_
								celdadatos(RSops("fdesp")) &_
								celdadatos(RSops("fentrada")) &_
								celdadatos(RSops("KPICTE")) &_
								celdadatos(RSops("KPIGRK")) &_
								celdadatos(RSops("FCG")) &_
								celdadatos(RSops("CG")) &_
								celdadatos(RSops("KPIADMIN")) &_
								celdadatos(RSops("facuse")) &_
								celdadatos(RSops("KPIACUSE"))
				If RSops("etapa") <> "" And IsNull(RSops("etapa"))= False  AND nivel <> 1 Then
					etapa = RSops("etapa")
					' Response.Write("referencia = " & referencia & "etapa consulta = " & RSops("etapa") & " etapa var = " & etapa & "contador = " & Cstr(cont) & "<br>")
					datos = datos & "</tr>" &_
								"<tr>" &_
									celdadatos("&nbsp;") &_
									celdaetapa(RSops("etapa") & " " & RSops("fetapa")) &_
								"</tr>"
					if nivel = 3 then
						datos = datos & "<tr>" &_
									celdadoble("&nbsp;") &_
									celdatriple(RSops("causa")) &_
									celdaobser(RSops("observacion"))
					end if
				End If
			Else
				' Response.Write("referencia = " & referencia & "etapa consulta = " & RSops("etapa") & " etapa var = " & etapa & "contador = " & Cstr(cont) & "<br>")
				if RSops("etapa") <> etapa and nivel <> 1 then
					etapa = RSops("etapa")
					datos = datos & celdadatos("&nbsp;") &_
									celdaetapa(RSops("etapa") & " " & RSops("fetapa")) &_
							"</tr>"
				End If
				if nivel = 3 then
					datos = datos & 	celdadoble("&nbsp;") &_
										celdatriple(RSops("causa")) &_
										celdaobser(RSops("observacion"))
				End If
			End If
			datos = datos & "</tr>"
			datos = Replace(datos,"<tr></tr>","")
			' Response.End
			RSops.MoveNext()
		Loop
		datos = datos & "</table>"
		html = ""
		' Response.End
		html = info & header & datos
	End If
End If
	
Function GeneraSQL
	SQL = ""
	SQL =			"SELECT i.refcia01 AS 'referencia', " &_
					"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pedimento', " &_
					"i.cvecli01 AS 'cvecli', " &_
					"i.nomcli01 AS 'nomcli', " &_
					"i.totbul01 AS 'bultos', " &_
					"IF(c.feta01 IS NULL OR c.feta01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.feta01,'%d-%m-%Y')) AS 'feta', " &_
					"IF(c.fdoc01 IS NULL OR c.fdoc01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.fdoc01,'%d-%m-%Y')) AS 'fdocs', " &_
					"IF(c.frev01 IS NULL OR c.frev01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.frev01,'%d-%m-%Y')) AS 'frev', " &_
					"IF(c.fpre01 IS NULL OR c.fpre01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.fpre01,'%d-%m-%Y')) AS 'fprev', " &_
					"IF(c.fdsp01 IS NULL OR c.fdsp01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.fdsp01,'%d-%m-%Y')) AS 'fdesp', " &_
					"IF(i.fecpag01 IS NULL OR i.fecpag01 = '0000-00-00', '&nbsp;', DATE_FORMAT(i.fecpag01,'%d-%m-%Y')) AS 'fpago', "
	if mov = "i" then
		SQL = SQL &	"IF(i.fecent01 IS NULL OR i.fecent01 = '0000-00-00' , '&nbsp;', DATE_FORMAT(i.fecent01, '%d-%m-%Y')) as 'fentrada', " &_
					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) - if( ((DAYOFWEEK(i.fecent01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(i.fecent01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(i.fecent01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) )))),(( TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) - if( ((DAYOFWEEK(i.fecent01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(i.fecent01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(i.fecent01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) * 2) ))))) as 'KPICTE', "
	else 
		SQL = SQL & "IF(c.frec01 IS NULL OR c.frec01 = '0000-00-00' , '&nbsp;', DATE_FORMAT(c.frec01, '%d-%m-%Y')) as 'fentrada', " &_
					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) - if( ((DAYOFWEEK(c.frec01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(c.frec01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(c.frec01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *1.5) )))),(( TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) - if( ((DAYOFWEEK(c.frec01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(c.frec01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(c.frec01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) * 2) ))))) as 'KPICTE', "
	end if
	SQL = SQL &		"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) - if( ((DAYOFWEEK(c.frev01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(c.frev01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(c.frev01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *1.5) )))), (( TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) - if( ((DAYOFWEEK(c.frev01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(c.frev01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(c.frev01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *2) ))))) as 'KPIGRK', " &_
					"IF(cta.fech31 IS NULL OR cta.fech31 = '0000-00-00','No Hay CG',DATE_FORMAT(MIN(cta.fech31),'%d-%m-%Y')) as 'FCG', " &_
					"IF(cta.cgas31 IS NULL OR cta.cgas31 = '', 'No Hay CG',MIN(cta.cgas31)) as 'CG', " &_
					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01) ) - if( ((DAYOFWEEK(c.fdsp01) -1) = 6 ), ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(c.fdsp01) -1) = 7, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(c.fdsp01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01)) )/ 7)) *1.5) )))), (( TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01) ) - if( ((DAYOFWEEK(c.fdsp01) -1) = 6 ), ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(c.fdsp01) -1) = 7, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(c.fdsp01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01)) )/ 7)) * 2) ))))) as 'KPIADMIN', " &_
					"IF(cta.frec31 IS NULL OR cta.frec31 = '0000-00-00', '&nbsp;', DATE_FORMAT(MIN(cta.frec31), '%d-%m-%Y')) as 'facuse', " &_
					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31)) ) - if( ((DAYOFWEEK(MIN(cta.fech31)) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(MIN(cta.fech31)) -1) = 7 , ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(MIN(cta.fech31)) -1)+TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31)) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *1.5) )))), (( TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31)) ) - if( ((DAYOFWEEK(MIN(cta.fech31)) -1) = 6 ), ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *2) - 1, if( (DAYOFWEEK(MIN(cta.fech31)) -1) = 7, ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(MIN(cta.fech31)) -1)+TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31)) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31))) )/ 7)) * 2) ))))) as 'KPIACUSE', " &_
					"ets.d_nombre AS 'etapa', " &_
					"etx.f_fecha AS 'fetapa', " &_
					"IF(caus.c01causa IS NULL OR caus.c01causa = '', 'Comentarios: ', caus.c01causa) AS 'causa', " &_
					"IF(etx.m_observ IS NULL OR etx.m_observ = '', '&nbsp;', etx.m_observ) AS 'observacion' " &_
					"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
					"LEFT JOIN " & strOficina & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 " &_
					"LEFT JOIN " & strOficina & "_extranet.d31refer as ctar ON ctar.refe31 = i.refcia01 " &_
					"LEFT JOIN " & strOficina & "_extranet.e31cgast as cta ON cta.cgas31 = ctar.cgas31 " &_
					"LEFT JOIN " & strOficina & "_status.etxpd AS etx ON etx.c_referencia = i.refcia01 " &_ 
					"LEFT JOIN " & strOficina & "_status.etaps AS ets ON ets.n_etapa = etx.n_etapa " &_
					"LEFT JOIN " & strOficina & "_status.c01caus AS caus ON caus.c01clavec = etx.clavec AND caus.c01rfc = i.rfccli01 " &_
					"WHERE i.firmae01 IS NOT NULL AND i.firmae01 <> '' AND i.cveped01 <> 'R1' AND c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "' " &_
					"AND (cta.esta31 <> 'C' or cta.esta31 IS NULL) AND (cta.fech31 >= c.fdsp01 OR cta.fech31 IS NULL) " &_
					filtroetapa &_
					filtrocliente &_
					"GROUP BY i.refcia01, etx.n_secuenc " &_
					"ORDER BY i.refcia01, etx.n_etapa "

	' Response.Write(SQL)
	' Response.End()
	GeneraSQL = SQL
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
	if texto = "" or IsNull(texto) = True Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function celdaetapa(texto)
	cell = 	"<td colspan=""21"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdaetapa = cell
end function

function celdaobser(texto)
	cell = 	"<td colspan=""17"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdaobser = cell
end function

function celdadoble(texto)
	cell = 	"<td colspan=""2"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadoble = cell
end function

function celdatriple(texto)
	cell = 	"<td colspan=""3"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdatriple = cell
end function
%>

<HTML>
	<HEAD>
		<TITLE>
			:: ....REPORTE DE SEGUIMIENTO DE OPERACIONES.... ::
		</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
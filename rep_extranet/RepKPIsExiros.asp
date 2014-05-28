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
		case "TOL"
			strOficina="tol"
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
	MesIn = UCase(MonthName(Month(fi)))
	DateI = Anioi & "/" & Mesi & "/" & Diai

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	MesFi = UCase(MonthName(Month(ff)))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	nocolumns = 0
	tablamov = ""
	if mov = "i" then
		movi = "IMPORTACION"
		tablamov = "ssdagi01"
		if strOficina="rku" then
			nocolumns = 22
		else
			nocolumns = 22
		end if
		query = GeneraSQL
	else
		movi = "EXPORTACION"
		tablamov = "ssdage01"
		if strOficina="rku" then
			nocolumns = 14
		else
			nocolumns = 14
		end if
		query = GeneraSQL
	end if
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write("DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	 'Response.Write(query)
	 'Response.End()
	 '
	Set RSops = CreateObject("ADODB.RecordSet")
	'response.write(query)
	'response.end()
	Set RSops = ConnStr.Execute(query)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
		if Tiporepo = 2 Then
			if mov = "i" Then 
				Response.Addheader "Content-Disposition", "attachment;filename=Reporte KPIs TAMSA-IMPO-.xls;"
			else
				Response.Addheader "Content-Disposition", "attachment;filename=Reporte KPIs TAMSA-EXPO-.xls;"
			End If
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<td colspan=""2"" rowspan=""4"">" &_
								"<img src = ""http://rkzego.no-ip.org/PortalMySQL/Extranet/ext-Images/LogoRepKPIExiros.png"" >" &_
								"</td>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"GRUPO REYES KURI, S.C" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>"
												if strOficina = "rku" then
													info = info & "VERACRUZ"  
												else 
													info = info & "" 
												end if
											info = info & "</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"DESPACHOS DE " & movi &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center> " &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
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
				"</table>"
		
		header = 	"<table    border = ""1"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#002060"" class = ""boton"">" &_
								celdahead("No.") &_
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Cliente") &_
								celdahead("Ejecutivo")
		if mov = "i" Then
			header = header & 	celdahead("Descripci&oacute;n de las Mercanc&iacute;as") &_
								celdahead("Pa&iacute;s Origen")
		Else
			header = header & 	celdahead("N&uacute;meros de Facturas")
		End If
		header = header & 		celdahead("Recepci&oacute;n de Documentos") &_
								celdahead("Entrada")
		if mov = "i" Then
			header = header &	celdahead("Revalidaci&oacute;n") &_
								celdahead("Previo")
		End If
		header = header &		celdahead("Pago") &_	
								celdahead("Despacho")
		if mov = "i" Then
			header = header &	celdahead("Tiempos de Despacho (Indicador GZ) (Revalidación vs Despacho) <= 2 días")
		End If
		
		if mov = "i" Then
			header = header &		celdahead("Tiempos de Despacho TAMSA (Entrada vs Despacho) <= 5 días Contenedores y 7 CS ")
		else
			header = header &		celdahead("Tiempos de Despacho indicador TAMSA (Entrada vs Despacho) <= 2 d&iacute;as")
		End If
		
		header = header &	celdahead("Observaciones Tr&aacute;fico")
		if mov = "i" Then
			header = header &	Celdahead("Causa del Desv&iacute;o (Arbol de Pérdidas)") &_
								celdahead("Total de Contenedores / Bultos") &_
								celdahead("Destino de la Mercanc&iacute;a")
		End If
		header = header &	celdahead("Fecha C.G.") &_
								celdahead("No. C.G.")
		if mov = "i" Then
			header = header & celdahead("Recepci&oacute;n de C.G.")
		End If
		header = header & 	celdahead("Tiempos  de Facturaci&oacute;n Despacho vs Recepci&oacute;n  C.G.") &_
								celdahead("Observaciones Administraci&oacute;n")
		
		header = header &	"</tr>"
		datos = ""
		Referencia = ""
		ubica = ""
		facturas = ""
		contenedores = ""
		total = ""
		iRenglon = 0
		Do Until RSops.EOF
			Referencia = RSops.Fields.Item("referencia").value
			if mov = "i" Then
				ubica = destinos(Referencia)
				total = totalconte(Referencia)
			Else
				facturas = contienefacturas(Referencia)
			End If
			obs = Observaciones(Referencia)
			iRenglon = iRenglon + 1
			datos = datos &	"<tr>" &_
								celdadatos(iRenglon) &_
								celdadatos(Referencia) &_
								celdadatos(RSops.Fields.Item("pedimento").Value) &_
								celdadatos(RSops.Fields.Item("nomcli").Value) &_
								celdadatos(RSops.Fields.Item("ejecutivo").Value)
			if mov = "i" Then
				datos = datos & celdadatos(RSops.Fields.Item("Descpro").Value)
			Else
				datos = datos & celdadatos(facturas)				
			End If
			if mov = "i" Then
				datos = datos & celdadatos(RSops.Fields.Item("porigen").Value)
			End If
			
			datos = datos & 	celdadatos(RSops.Fields.Item("fdocs").Value)
			datos = datos &		celdadatos(RSops.Fields.Item("fentrada").Value)
			if mov = "i" Then
				datos = datos &	celdadatos(RSops.Fields.Item("frev").Value) &_
								celdadatos(RSops.Fields.Item("fprev").Value)
			End If
			datos = datos & 	celdadatos(RSops.Fields.Item("fpago").Value) &_
								celdadatos(RSops.Fields.Item("fdesp").Value)
			if mov = "i" Then
				datos = datos & celdadatos(RSops.Fields.Item("KPIGRK").Value)
			end if
			datos = datos & 	celdadatos(RSops.Fields.Item("KPICTE").Value)
			datos = datos & celdadatos("")
			
			if mov = "i" Then
				datos = datos &	celdadatos(Causales(Referencia, "T",RSops.Fields.Item("rfccli01").Value)) &_
						celdadatos(RSops.Fields.Item("totbulx").value) & _
						celdadatos(ubica)
			End If
			datos = datos &	celdadatos(RSops.Fields.Item("FCG").Value) &_
							celdadatos(RSops.Fields.Item("CG").Value)
			if mov = "i" Then
				datos = datos &	celdadatos(RSops.Fields.Item("facuse").Value)
			End If
				datos = datos &	celdadatos(RSops.Fields.Item("KPIADMIN").Value) &_
						celdadatos("&nbsp;")
			
			datos = datos &	"</tr>"
			Rsops.MoveNext()
		Loop
	
	prom = ""
	html = info & header & datos & "</table><br>"
	
	
	End If
end if


function celdahead(texto)
	cell = "<td bgcolor = ""#002060"" height = ""88"" align=""center"" width=""150"" style=""vertical-align:middle"" >" &_
				"<center>" &_
					"<strong>" &_
						"<font color=""#FFFFFF"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
							"<i>" &_
								texto &_
							"</i>" &_
						"</font>" &_
					"</strong>" &_
				"</center>" &_
			"</td>"
	celdahead = cell
end function

function celdadatos(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"" height = ""50"" style=""vertical-align:middle"">" &_
				"<font size=""2"" face=""Arial"" >" &_
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
	SQL = 	"SELECT  i.refcia01 AS 'referencia', " &_
			"i.cveped01 as 'Clvp', i.rfccli01, " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pedimento', " &_
			"i.cvecli01 AS 'cvecli', " &_
			"i.nomcli01 AS 'nomcli', " &_
			"group_concat(distinct fr.d_mer102) AS 'Descpro2', " &_
			"( select group_concat(distinct art.desc05) from " & strOficina & "_extranet.d05artic as art where art.refe05 = i.refcia01) AS 'Descpro', " &_
			" i.nompro01 as Nomproveedor, "&_
			"( select group_concat(distinct art.pedi05 SEPARATOR '/ ') from " & strOficina & "_extranet.d05artic as art where art.refe05 = i.refcia01) AS 'OC', " &_
			"( select group_concat(distinct art.tpmerc05) from " & strOficina & "_extranet.d05artic as art where art.refe05 = i.refcia01) AS 'TM', " &_
			"'jajaj' AS 'DescproX', " &_
			"fr.paiori02 AS 'porigen', " &_
			"i.totbul01 AS 'bultos', " &_
			"IF(c.feta01 IS NULL OR c.feta01 = '0000-00-00', '', DATE_FORMAT(c.feta01,'%d-%m-%Y')) AS 'feta',  " &_
			"IF(c.fdoc01 IS NULL OR c.fdoc01 = '0000-00-00', '', DATE_FORMAT(c.fdoc01,'%d-%m-%Y')) AS 'fdocs',  " &_
			"IF(c.frev01 IS NULL OR c.frev01 = '0000-00-00', '', DATE_FORMAT(c.frev01,'%d-%m-%Y')) AS 'frev',  " &_
			"IF(c.fpre01 IS NULL OR c.fpre01 = '0000-00-00', if((select ip.cveide11  from " & stroficina & "_extranet.ssiped11 as ip where ip.refcia11 =i.refcia01 and ip.cveide11 ='RO')='RO',DATE_FORMAT(i.fecpag01,'%d-%m-%Y'), ''), DATE_FORMAT(c.fpre01,'%d-%m-%Y')) AS 'fprev',  " &_
			"DATE_FORMAT(c.fdsp01,'%d-%m-%Y') AS 'fdesp' ,  " &_
			"DATE_FORMAT(MIN(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d/%m/%Y') AS 'frobot', " &_
			"IF(i.fecpag01 IS NULL OR i.fecpag01 = '0000-00-00', '', DATE_FORMAT(i.fecpag01,'%d-%m-%Y')) AS 'fpago', "
			if mov = "i" Then
				SQL = SQL & "IF(i.fecent01 IS NULL OR i.fecent01 = '0000-00-00', '', DATE_FORMAT(i.fecent01,'%d-%m-%Y')) AS 'fentrada', " &_
							kpi("","i.fecent01", "c.fdsp01") & "as KPICTE, "
			Else
				SQL = SQL & "IF(i.fecpre01 IS NULL OR i.fecpre01 = '0000-00-00' , '', DATE_FORMAT(i.fecpre01, '%d-%m-%Y')) AS 'fentrada', " &_
							kpi("","i.fecpre01", "c.fdsp01") & "as KPICTE, "
			end if
	SQL = SQL & kpi("", "c.frev01", "c.fdsp01") & "AS 'KPIGRK', " &_
			"IF(cta.fech31 IS NULL or  cta.fech31= '0000-00-00','No Hay CG',DATE_FORMAT(cta.fech31,'%d-%m-%Y')) as 'FCGx',  " &_
			" (select DATE_FORMAT(max( x2.fech31), '%d-%m-%Y') from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as FCG, " & _
			" (select group_concat(distinct x2.fech31) from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as FCG33, " & _
			"@maxfcg:= (select max( x2.fech31) from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as maxFCG, " & _
			"IF(cta.cgas31 IS NULL OR cta.cgas31 = '', 'No se ha Facturado',cta.cgas31) as 'CGx',  " &_
			" (select group_concat(distinct x2.cgas31) from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as CG, " & _
			kpi("", "c.fdsp01", "@maxfcg") & " AS 'KPIADMIN', " &_
			"IF(cta.frec31 IS NULL OR cta.frec31 = '0000-00-00', '', DATE_FORMAT(cta.frec31, '%d-%m-%Y')) as 'facusex', " &_
			" (select if(max(x2.frec31) is not null and max(x2.frec31) <> '0000-00-0', DATE_FORMAT(max(x2.frec31), '%d-%m-%Y'),'') from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as 'facuse', " & _
			" (select group_concat(distinct x2.frec31) from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as facuse33, " & _
			" @maxfacu:=(select max(x2.frec31) from  " & strOficina & "_extranet.e31cgast as x2 inner join  " & strOficina & "_extranet.d31refer as x3 on x3.cgas31 = x2.cgas31 where x2.esta31 <>'C' and x3.refe31 = i.refcia01) as 'Maxfacuse', " & _
			kpi("", "@maxfcg", "@maxfacu") & "AS 'KPIACUSE', " &_
			"DATE_FORMAT(MAX(c.fcot01), '%d-%m-%Y') AS 'fcotizacion', " &_
			"DATE_FORMAT(MAX(d11.fech11), '%d-%m-%Y') AS 'fdeposito', " &_
			" ( select concat(sum(x1.cant01),' (',group_concat(distinct  x1.clas01),')') from " & strOficina & "_extranet.d01conte as x1 where x1.refe01 =i.refcia01) as totbulx,  " & _
			"if(c.ejecli01 =0,'',d18.ejec18) AS 'ejecutivo' ," &_
			"(select group_concat(distinct g.numgui04 separator '/ ')from " & strOficina & "_extranet.ssguia04 as g where g.refcia04=i.refcia01 and g.adusec04=i.adusec01 and g.patent04=i.patent01 and g.idngui04=1) AS 'GUIABL' " &_
			"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
			"LEFT JOIN " & strOficina & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 " &_
			"LEFT JOIN " & strOficina & "_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 " &_
			"LEFT JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
			"LEFT JOIN " & strOficina & "_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 " &_
			"LEFT JOIN " & strOficina & "_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 and cta.esta31 <> 'C' " &_
				"AND (cta.esta31= 'I' ) " &_
			"LEFT JOIN " & strOficina & "_extranet.d11movim AS d11 ON d11.refe11 = i.refcia01 AND d11.conc11 = 'ANT' " &_
			"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 and bs.Numpat01=i.patent01 and bs.Detsit01 ='730'  " & _
			"WHERE i.firmae01 IS NOT NULL AND i.firmae01 <> '' AND i.cveped01 <> 'R1' " &_
			"AND ((c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "') AND c.fdsp01 <> '00-00-0000') " & condicion &_ 
			"GROUP BY i.refcia01 " &_
			"ORDER BY c.fdsp01" ' "ORDER BY i.refcia01"

			'"AND ((c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "') AND c.fdsp01 <> '00-00-0000') " & condicion &_ '28-08-2012
			
			'"i " &_
			'"LEFT JOIN trackingbahia.bit_soia as bs ON bs.frmsaai01 = i.refcia01 AND bs.Numped01 = i.numped01 AND bs.Adusec01 = i.adusec01 AND bs.rfccli01 = i.rfccli01 AND bs.Numpat01 = i.patent01 " &_
			
			
			
			'"AND (cta.esta31 <> 'C' or cta.esta31 IS NULL) " &_
			'	"AND (cta.fech31 >= c.fdsp01 Or cta.fech31 IS NULL) " &_
			
	' Response.Write(SQL)
	'Response.End
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
			"FROM " & stroficina & "_extranet.d01conte AS d01 " &_
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
				"FROM " & stroficina & "_status.etxpd " &_
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

Function Causales(refe, tipo,cliente)
	causas =	""
	SQLCausales = 	""
	SQLCausales = 	"SELECT DISTINCT etx.c_referencia, cau.c01causa, cau.c01tipoc " &_
					"FROM "&strOficina&"_status.etxpd AS etx " &_
					"INNER JOIN "&strOficina&"_status.c01caus AS cau ON cau.c01clavec = etx.clavec " &_
					"WHERE etx.c_referencia = '" & refe & "' AND cau.c01causa <> '' AND cau.c01tipoc LIKE '" & tipo & "' "'and cau.c01tipoo<>0  "'and  cau.c01rfc ='"&cliente&"' "

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
		<TITLE>::.... REPORTE DE SEGUIMIENTO DE OPERACIONES .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=150000
	 On Error Resume Next
strTipoUsuario = request.Form("TipoUser")

strPermisos = Request.Form("Permisos")

permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

if not permi = "" then
	permi = "  and (" & permi & ") "
end if



cve=request.form("cve")
mov=request.form("mov")
fi=trim(request.form("fi"))
ff=trim(request.form("ff"))

if fi <> "" and IsNull(fi) = False Then
	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & Diai
End If
if ff <> "" and IsNull(ff) = False Then
	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
End If
Vrfc = Request.Form("rfcCliente")
if strTipoUsuario = 004 then
	Vckcve = 1
	Vclave = "Todos"
else
Vckcve = Request.Form("ckcve")
Vclave = Request.Form("txtCliente")
end if
Tiporepo = Request.Form("TipRep")
pedi = Request.Form("pedi")
refcia = Request.Form("refe")
artic = Request.Form("artic")
TipoMer = Request.Form("TipMer")

'Response.Write("Vckcve = " & Vckcve & "<br> Vclave = " & vclave )
'Response.Write(permi)
'Response.End

filtrowhere = ""
if Datei <> "" and IsNull(Datei) = False Then
	filtrowhere = filtrowhere & "AND i.fecpag01 >= '" & Datei & "' "
End If
If Datef <> "" and IsNull(Datef) = False Then 
	filtrowhere = filtrowhere & "AND i.fecpag01 <= '" & Datef & "' "
End If
if pedi <> "" and IsNull(pedi) = False Then
	filtrowhere = filtrowhere & "AND CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) like '%" & pedi & "%' "
End If

if refcia <> "" And IsNull(refcia) = False Then
	filtrowhere = filtrowhere & "AND i.refcia01 like '" & refcia & "' "
End If

if artic <> "" And IsNull(artic) = False Then
	filtrowhere = filtrowhere & "AND d05.cpro05 like '%" & artic & "%' "
End If

if TipoMer = "D" then
	filtrowhere = filtrowhere & "AND d05.tpmerc05 = 'PT' "
else
	if TipMer = "R" then
		filtrowhere = filtrowhere & "AND d05.tpmerc05 <> 'PT' "
	End If
End If

If Vckcve = 0 Then 
	filtrocliente = "AND i.rfccli01 = '" & Vrfc & "' "
Else 
	If Vclave = "Todos" Then 
		filtrocliente = permi
		filtrocliente = Replace(filtrocliente, "and (cvecli01=0 )", "")
	Else
		filtrocliente = "AND i.cvecli01 = " & Vclave & " "
	End If
End If

filtrowhere = filtrowhere & filtrocliente
 'Response.Write(filtrowhere)
 'Response.End

'filtrowhere = "AND i.refcia01 = 'RKU10-06206' OR i.refcia01 = 'RKU10-06382' or i.refcia01 = 'RKU10-00153A' "

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
	'Response.Write(query)
	'Response.End()
	Set ConnStr = Server.CreateObject("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = Server.CreateObject("ADODB.Recordset")
	Set RSops = ConnStr.Execute(query)
	if RSops.Eof = True and RSops.Bof = True Then
	'if(query<>"oki") then
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
		'RSops.MoveFirst
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
						"<tr>" &_
							"<strong>" &_
								"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
									"<td colspan=""59"">" &_
										"<p align=""left"">" &_
											"::.... REPORTE DE OPERACIONES TEMPORALES.... ::" &_
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
								celdahead("Numero de Pedimento") &_
								celdahead("Tipo de Operacion") &_
								celdahead("Clave de Pedimento") &_
								celdahead("Regimen") &_
								celdahead("Tipo de Cambio") &_
								celdahead("Aduana") &_
								celdahead("Medio de Transportes") &_
								celdahead("Valor Dolares Pedimento") &_
								celdahead("Valor Aduana Pedimento")
								
							if mov = "i" then
								header = header & celdahead("Valor Comercial Pedimento")
							end if
							
								header = header & celdahead("Importador") &_
										 celdahead("Valor Seguros")
								
							if mov = "i" then
								header = header & celdahead("Seguros") &_
										 celdahead("Fletes") &_
										 celdahead("Embalajes")
							end if
							
								header = header & celdahead("Otros Incrementables")
								
							if mov = "i" then
								header = header & celdahead("Fecha Entrada")
							end if
								
								header = header & celdahead("Fecha Pago") &_
								celdahead("DTA") &_
								celdahead("IGIE") &_
								celdahead("IVA") &_
								celdahead("PREV") &_
								celdahead("No. Pedimento Al Que Rectifica") &_
								celdahead("Fecha Pago") &_
								celdahead("No. Pedimento Descargo") &_
								celdahead("Fecha Pago") &_
								celdahead("Regimen") &_
								celdahead("Num. Factura y Fecha") &_
								celdahead("Proveedor") &_
								celdahead("INCOTERM") &_
								celdahead("Moneda Factura") &_
								celdahead("Val. Mon. Factura") &_
								celdahead("Factor Mon. Factura") &_
								celdahead("Valor Dolares Factura") &_
								celdahead("Transporte Identificacion") &_
								celdahead("No. Guia/BL/ID") &_
								celdahead("Numero/Tipo/Contenedor") &_
								celdahead("Fraccion") &_
								celdahead("Orden Fraccion") &_
								celdahead("UMC") &_
								celdahead("Cantidad UMC") &_
								celdahead("Pais V/C") &_
								celdahead("Pais O/D") &_
								celdahead("Descripcion Mercancia") &_
								celdahead("Valor Aduana Fraccion") &_
								celdahead("Imp. Precio Pagado") &_
								celdahead("Precio Unitario Mon. Ext.") &_
								celdahead("Precio Unitario Dolares") &_
								celdahead("Precio Unitario Mon. Nac.") &_
								celdahead("Orden De Compra") &_
								celdahead("Item Bann") &_
								celdahead("Destino") &_
								celdahead("Tipo Mercancia") &_
								celdahead("Identificador") &_
								celdahead("Cuenta de Gastos") &_
								celdahead("Total CG") &_
								celdahead("Fecha Vencimiento") &_
								celdahead("Dias Restantes") &_
								celdahead("Destino Final") &_
							"</tr>"
		' Response.Write(info & header)
		' Response.End
		datos = ""
		referencia = ""
		anterior = ""
		recti = ""
		fecharecti = ""
		Set RSrec = Server.CreateObject("ADODB.Recordset")
		Set RSdes = Server.CreateObject("ADODB.Recordset")
		
		'Do Until RSops.Eof	
		while not RSops.eof
				
			if referencia <> RSops("referencia") Then
				
				referencia = RSops("referencia")
				cveped = RSops("claveped")
				
				if cveped = "R1" then
					query = obtieneA1(referencia)
				
					' response.write(referencia)
					' response.end()
				
					Set RSrec = ConnStr.Execute(query)
					if RSrec.Eof = True and RSrec.Bof = True Then
						recti = "&nbsp;"
						fecharecti = "&nbsp;"
					' --Response.Write("NO HUBO RECTIFICACION refe = " & referencia & "<br>")
					' --Response.End()
					Else
						recti = RSrec("pediori")
						fecharecti = RSrec("fpagori")
					' --Response.Write("hubo recti refe = " & referencia & "<br>")
					' --Response.End()
					End If
				else
					recti = "&nbsp;"
					fecharecti = "&nbsp;"
				end if
				
				query = ObtieneDesc(referencia)
								'--response.write(query)
				'--response.end()
				Set RSdes = ConnStr.Execute(query)
				
				If RSdes.Eof = True and RSdes.Bof = True Then
				    pedesc = "&nbsp;"
					regidesc = "&nbsp;"
					fechadesc = "&nbsp;"
					' --Response.Write("NO hubo descargo refe = " & referencia & "<br>")
				Else
					
					pedidesc = RSdes("pedidesc")
					regidesc = RSdes("regidesc")
					fechadesc = RSdes("fechadesc")
					' --Response.Write("hubo descargo refe = " & referencia & "<br>")
				End If
				
			end if
			datos = datos & "<tr>" &_
								celdadatos(RSops("pedimento")) &_
								celdadatos(RSops("toper")) &_
								celdadatos(RSops("claveped")) &_
								celdadatos(RSops("regimen")) &_
								celdadatos(RSops("tipocam")) &_
								celdadatos(RSops("aduana")) &_
								celdadatos(RSops("mtransporte")) &_
								celdadatos(RSops("vdolares")) &_
								celdadatos(RSops("vaduana"))
								
							if mov = "i" then
								datos = datos & celdadatos(RSops("vcomercial"))
							end if
							
								datos = datos & celdadatos(RSops("cliente")) &_
								celdadatos(RSops("valseguros"))
								
							if mov = "i" then
								datos = datos & celdadatos(RSops("seguros")) &_
								celdadatos(RSops("fletes")) &_
								celdadatos(RSops("embalajes"))
							end if
							
								datos = datos & celdadatos(RSops("incrementables"))
								
							if mov = "i" then
								datos = datos & celdadatos(RSops("fentrada"))
							end if
							
								datos = datos & celdadatos(RSops("fpago")) &_
								celdadatos(RSops("dta")) &_
								celdadatos(RSops("adv")) &_
								celdadatos(RSops("iva")) &_
								celdadatos(RSops("prev")) &_
								celdadatos(recti) &_
								celdadatos(fecharecti) &_
								celdadatos(pedidesc) &_
								celdadatos(regidesc) &_
								celdadatos(fechadesc) &_
								celdadatos(RSops("factura")) &_
								celdadatos(RSops("proveedor")) &_
								celdadatos(RSops("incoterm")) &_
								celdadatos(RSops("monedafac")) &_
								celdadatos(RSops("vmonedaext")) &_
								celdadatos(RSops("factormon")) &_
								celdadatos(RSops("vdolares")) &_
								celdadatos(RSops("Barco")) &_
								celdadatos(RSops("guia")) &_
								celdadatos(RSops("conte")) &_
								celdadatos(RSops("fraccion")) &_
								celdadatos(RSops("orden")) &_
								celdadatos(RSops("umc")) &_
								celdadatos(RSops("cantidadumc")) &_
								celdadatos(RSops("paiscv")) &_
								celdadatos(RSops("paisod")) &_
								celdadatos(RSops.Fields.Item("descripcionmerc").Value) &_
								celdadatos(RSops("vaduanfrac")) &_
								celdadatos(RSops("impprecio")) &_
								celdadatos(RSops("preciounimext")) &_
								celdadatos(RSops("preciounidlls")) &_
								celdadatos(RSops("preciounimn")) &_
								celdadatos(RSops("ordencompra")) &_
								celdadatos(RSops("itembann")) &_
								celdadatos(RSops.Fields.Item("observacion").Value) &_
								celdadatos(RSops("tpmerc")) &_
								celdadatos(RSops("ident")) &_
								celdadatos(RSops("cg")) &_
								celdadatos(RSops("totalcg")) &_
								celdadatos(RSops("tiempo")) &_
								celdadatos(RSops("dias")) &_
								celdadatos(RSops("desfi")) &_
							"</tr>"
			RSops.MoveNext()
		wend 
		Set RSrec = Nothing
		datos = datos & "</table>"
		html = ""
		' Response.End
		html = info & header & datos
	End If
End If
	
Function GeneraSQL
	SQL = 			""
	SQL =			"SELECT i.refcia01 as 'referencia', " &_
					"CONCAT_WS('-', i.patent01, i.numped01) AS 'pedimento', " &_
					"i.cveped01 AS 'claveped', " &_
					"i.regime01 AS 'regimen', " &_
					"i.tipcam01 AS 'tipocam', " &_
					"i.adusec01 AS 'aduana', " &_
					"mtr.descri30 AS 'mtransporte', " &_
					"i.valdol01 AS 'vdolares', " &_
					"ROUND((i.tipcam01 * i.valdol01),0) AS 'vaduana', " &_
					"i.nomcli01 AS 'cliente', " &_
					"i.valseg01 AS 'valseguros', "
	if mov = "i" then 
		SQL = SQL &	"i.valmer01 AS 'vcomercial', " &_
					"'IMP' AS 'toper', " &_
					"i.rsegur01 AS 'seguros', " &_
					"i.rflete01 AS 'fletes', " &_
					"i.rembal01 AS 'embalajes', " &_
					"DATE_FORMAT(i.fecent01, '%d/%m/%Y') AS 'fentrada', " 
	Else
		SQL = SQL &	"'EXP' AS 'toper', "
	End If
	SQL = SQL &		"i.incble01 AS 'incrementables', " &_
					"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpago', " &_
					"IFNULL(dta.import36, 0) AS 'dta', " &_
					"IFNULL(adv.import36, 0) AS 'adv', " &_
					"IFNULL(iva.import36, 0) AS 'iva', " &_
					"IFNULL(pre.import36, 0) AS 'prev', " &_
					"CONCAT_WS('-', i.nompro01, i.taxpro01) AS 'proveedor', " &_
					"CONCAT_WS('--', fac.numfac39, CAST(fac.fecfac39 AS CHAR)) AS 'factura', " &_
					"fac.terfac39 AS 'INCOTERM', " &_
					"fac.monfac39 AS 'monedafac', " &_
					"fac.valmex39 AS 'vmonedaext', " &_
					"fac.facmon39 AS 'factormon', " &_
					"fac.valdls39 AS 'vdolares', " &_
					"reb.nombar17 AS 'barco', " &_
					"gui.numgui04 AS 'guia', " &_
					"con.numcon40 AS 'conte', " &_
					"fr.fraarn02 AS 'fraccion', " &_
					"fr.ordfra02 AS 'orden', " &_
					"mcom.descri31 AS 'umc', " &_
					"fr.cancom02 AS 'cantidadumc', " &_
					"fr.paiscv02 AS 'paiscv', " &_
					"fr.paiori02 AS 'paisod', " &_
					"REPLACE(REPLACE(REPLACE(d05.desc05,'\n',' '),'\r',' '),'\a',' ') AS 'descripcionmerc', " &_
					"d05.vafa05 AS 'impprecio', " &_
					"fr.vaduan02 AS 'vaduanfrac', " &_
					"d05.pedi05 AS 'ordencompra', " &_
					"ROUND((d05.vafa05 / d05.caco05),2) AS 'preciounimext', " &_
					"ROUND(((d05.vafa05 / d05.caco05) * i.factmo01),2) AS 'preciounidlls', " &_
					"ROUND(((d05.vafa05 / d05.caco05) * i.factmo01 * i.tipcam01),2) AS 'preciounimn', " &_
					"d05.cpro05 AS 'itembann', " &_
					"REPLACE(REPLACE(REPLACE(REPLACE(d05.obse05, ';', ''), '\n', ''), '\r', ''),'\a','') AS 'observacion', " &_
					"d05.tpmerc05 AS 'tpmerc', " &_
					"DATE_ADD(i.fecpag01, INTERVAL 10 YEAR) AS 'tiempo', " &_
					"DATEDIFF(DATE_ADD(i.fecpag01, INTERVAL 10 YEAR), CURRENT_DATE()) AS 'dias', " &_
					"CONCAT_WS('  ', par.cveide12, par.comide12) AS 'ident', " &_
					"e31.cgas31 AS 'CG', " &_
					"e31.tota31 AS 'totalcg', " &_
					"des.nomb07 AS 'desfi' " &_
					"FROM " & StrOficina & "_extranet." & tablamov & " AS i " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssmtra30 AS mtr ON mtr.clavet30 = i.cvemtr01 " &_
					"LEFT JOIN " & StrOficina & "_extranet.sscont40 AS con ON con.refcia40 = i.refcia01 " &_
					"LEFT JOIN " & StrOficina & "_extranet.d01conte AS cte ON cte.refe01 = i.refcia01 " &_
					"LEFT JOIN " & StrOficina & "_extranet.c07desti AS des ON des.cdes07 = cte.cdes01 " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssreba17 AS reb ON reb.regbar17 = i.regbar01 " &_
					"LEFT JOIN " & StrOficina & "_extranet.sscont36 AS dta ON dta.refcia36 = i.refcia01 AND dta.cveimp36 = 1 " &_
					"LEFT JOIN " & StrOficina & "_extranet.sscont36 AS adv ON adv.refcia36 = i.refcia01 AND adv.cveimp36 = 6 " &_
					"LEFT JOIN " & StrOficina & "_extranet.sscont36 AS iva ON iva.refcia36 = i.refcia01 AND iva.cveimp36 = 3 " &_
					"LEFT JOIN " & StrOficina & "_extranet.sscont36 AS pre ON pre.refcia36 = i.refcia01 AND pre.cveimp36 = 15 " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssfact39 AS fac ON fac.refcia39 = i.refcia01 AND fac.adusec39 = i.adusec01 AND i.patent01 = fac.patent39 " &_
					"LEFT JOIN " & StrOficina & "_extranet.d05artic AS d05 ON d05.refe05 = i.refcia01 AND fac.numfac39 = d05.fact05 " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssfrac02 AS fr ON fr.refcia02 = i.refcia01 AND fr.adusec02 = i.adusec01 AND i.patent01 = fr.patent02 " &_
						"AND fr.fraarn02 = d05.frac05 AND fr.ordfra02 = d05.agru05 " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssguia04 AS gui ON gui.refcia04 = i.refcia01 AND i.patent01 = gui.patent04 " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssumed31 AS mcom ON mcom.clavem31 = fr.u_medc02 " &_
					"LEFT JOIN " & StrOficina & "_extranet.ssipar12 AS par ON i.refcia01 = par.refcia12 AND d05.agru05 = par.ordfra12 AND fr.ordfra02 = par.ordfra12 " &_
					"LEFT JOIN " & StrOficina & "_extranet.d31refer AS d31 ON d31.refe31 = i.refcia01 " &_
					"LEFT JOIN " & StrOficina & "_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
					"WHERE i.firmae01 <> '' AND i.firmae01 IS NOT NULL " & filtrowhere &_
					"GROUP BY i.refcia01,fac.numfac39,d05.item05,fr.fraarn02,fr.ordfra02,d05.pfac05 " &_
					"ORDER BY referencia, fr.ordfra02 "
					'&_	"LIMIT 500 "
					
	' Response.Write(SQL)
	' Response.End()
	GeneraSQL = SQL
End Function

function celdahead(texto)
	 'On Error Resume Next
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
	 On Error Resume Next
	if texto = "" Or IsNull(texto) = True Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"">" &_
				"<font size=""1"" face=""Arial"">" &_
					cstr(texto) &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

Function obtieneA1(refe)
	if refe = "" Or IsNull(refe) = True Then
		refe = "nothing"
	End If
	' Refe = 	"DAI10-10835A"
	SQL = 	"SELECT i.refcia01 AS 'reforg', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pediori', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpagori' " &_
			"FROM rku_extranet.ssrecp06 AS rec " &_
			"INNER JOIN rku_extranet." & tablamov & " AS i ON i.refcia01 = rec.reforg06 " &_
			"WHERE rec.refcia06 = '" & refe & "' AND i.firmae01 <> '' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'reforg', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pediori', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpagori' " &_
			"FROM dai_extranet.ssrecp06 AS rec " &_
			"INNER JOIN dai_extranet." & tablamov & " AS i ON i.refcia01 = rec.reforg06 " &_
			"WHERE rec.refcia06 = '" & refe & "' AND i.firmae01 <> '' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'reforg', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pediori', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpagori' " &_
			"FROM sap_extranet.ssrecp06 AS rec " &_
			"INNER JOIN sap_extranet." & tablamov & " AS i ON i.refcia01 = rec.reforg06 " &_
			"WHERE rec.refcia06 = '" & refe & "' AND i.firmae01 <> '' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'reforg', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pediori', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpagori' " &_
			"FROM lzr_extranet.ssrecp06 AS rec " &_
			"INNER JOIN lzr_extranet." & tablamov & " AS i ON i.refcia01 = rec.reforg06 " &_
			"WHERE rec.refcia06 = '" & refe & "' AND i.firmae01 <> '' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'reforg', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pediori', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpagori' " &_
			"FROM ceg_extranet.ssrecp06 AS rec " &_
			"INNER JOIN ceg_extranet." & tablamov & " AS i ON i.refcia01 = rec.reforg06 " &_
			"WHERE rec.refcia06 = '" & refe & "' AND i.firmae01 <> '' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'reforg', " &_
			"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pediori', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fpagori' " &_
			"FROM tol_extranet.ssrecp06 AS rec " &_
			"INNER JOIN tol_extranet." & tablamov & " AS i ON i.refcia01 = rec.reforg06 " &_
			"WHERE rec.refcia06 = '" & refe & "' AND i.firmae01 <> '' "
	' Response.Write(SQL)
	' Response.End()
	obtieneA1 = SQL
End Function

Function ObtieneDesc(refe)
	if refe = "" Or IsNull(refe) = True Then
		refe = "nothing"
	End If
	' refe = "DAI10-3751"
	SQL = ""
	SQL = 	"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM rku_extranet.ssdesc05 AS des " &_
			"INNER JOIN rku_extranet.ssdagi01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM dai_extranet.ssdesc05 AS des " &_
			"INNER JOIN dai_extranet.ssdagi01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM sap_extranet.ssdesc05 AS des " &_
			"INNER JOIN sap_extranet.ssdagi01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM lzr_extranet.ssdesc05 AS des " &_
			"INNER JOIN lzr_extranet.ssdagi01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM ceg_extranet.ssdesc05 AS des " &_
			"INNER JOIN ceg_extranet.ssdagi01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM tol_extranet.ssdesc05 AS des " &_
			"INNER JOIN tol_extranet.ssdagi01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM rku_extranet.ssdesc05 AS des " &_
			"INNER JOIN rku_extranet.ssdage01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM dai_extranet.ssdesc05 AS des " &_
			"INNER JOIN dai_extranet.ssdage01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM sap_extranet.ssdesc05 AS des " &_
			"INNER JOIN sap_extranet.ssdage01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM lzr_extranet.ssdesc05 AS des " &_
			"INNER JOIN lzr_extranet.ssdage01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM ceg_extranet.ssdesc05 AS des " &_
			"INNER JOIN ceg_extranet.ssdage01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' " &_
			"UNION ALL " &_
			"SELECT i.refcia01 AS 'referencia', " &_
			"CONCAT_WS('-', CONCAT(des.cveado05, des.cveseo05), des.cveago05, des.docori05) as 'pedidesc', " &_
			"DATE_FORMAT(i.fecpag01, '%d/%m/%Y') AS 'fechadesc', " &_
			"i.regime01 AS 'regidesc' " &_
			"FROM tol_extranet.ssdesc05 AS des " &_
			"INNER JOIN tol_extranet.ssdage01 AS i ON i.refcia01 = des.refcia05 " &_
			"where des.refcia05 = '" & refe & "' "
	' Response.Write(SQL)
	' Response.End()
	ObtieneDesc = SQL
End Function
%>

<HTML>
	<HEAD>
		<TITLE>
			:: ....REPORTE DE OPERACIONES TEMPORALES.... ::
		</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000
On Error Resume Next
'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/Rep_Anexo24-03122012-asv.asp

' strTipoUsuario = request.Form("TipoUser")
' strPermisos = Request.Form("Permisos")
' permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
 strOficina="sap"'Request.Form("OficinaG")
' strCortaR=Request.Form("cortaR")
	RFCliente="SVD000317AH4"'
	'RFCliente="JIG101229FM4"
	fi="2013-12-01"'trim(request.form("fi"))
	ff="2013-12-31" 'trim(request.form("ff"))
	' Vrfc=Request.Form("rfcCliente")
	' bclientes=Request.Form("Enviar")


	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & DiaI

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	
' if not permi = "" then
	' permi = "  and (" & permi & ") "
' end if
' AplicaFiltro = False
' strFiltroCliente = ""
' strFiltroCliente = request.Form("txtCliente")
 mov="a"'request.form("mov")

' Tiporepo = Request.Form("TipRep")

' if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	' blnAplicaFiltro = true
' end if
' if blnAplicaFiltro then
	' permi = " AND cvecli01 =" & strFiltroCliente
' end if
' if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	' permi = ""
' end if

' if  Session("GAduana") = "" then
	' html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
' else

	' if mov = "i" then
		' movi = "IMPORTACION "
		' query = GeneraSQL(mov)
	' elseif mov="e" then
		' movi="EXPORTACION "
		' query=GeneraSQL(mov)
	' elseif mov="a" then
		 movi = "IMPORTACION / EXPORTACION"
		 query = GeneraSQL(mov)

	' end if
	' 'response.write(query&strOficina)
	' 'response.end()
	 nocolumns = 16
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	
	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
		
		'if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;filename=Rep_Continental_"&DiaI&"-"&Mesi&"_"&DiaF&"-"&MesF&"-"&strOficina&".xls"
			Response.ContentType = "application/vnd.ms-excel"
		'End If
		info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td  align=""center"" colspan=""" & nocolumns & """></font></p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p><font color=""#000000"" size=""4"" face=""Arial, Helvetica, sans-serif"">: : Continental : :</font></p>" &_
											"<p><font color=""#000000"" size=""2"" face=""Arial, Helvetica, sans-serif"">"&movi&" DEL "&DateI&" al "&DateF&"</font>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("Referencia","#A11212") &_
								celdahead("Pedimento","#A11212") &_
								celdahead("Fecha de pago","#A11212") &_
								celdahead("Fecha de Entrada","#A11212")&_
								celdahead("Tipo de cambio","#A11212") &_
								celdahead("Regimen","#A11212") &_
								celdahead("Clave Pedimento","#A11212") &_
								celdahead("Patente","#A11212") &_
								celdahead("Aduana","#A11212") &_
								celdahead("Seccion","#A11212") &_
								celdahead("Valor Seguros","#A11212") &_
								celdahead("Seguros","#A11212") &_
								celdahead("Fletes","#A11212") &_
								celdahead("Embalajes","#A11212") &_
								celdahead("Otros  incrementables","#A11212") &_
								celdahead("DTA","#A11212") &_
								celdahead("Adicional","#A11212") &_
								celdahead("B/L","#A11212") &_
								celdahead("Buque","#A11212") &_
								celdahead("Descripcion Fraccion","#A11212") &_
								celdahead("Factura","#A11212") &_
								celdahead("Fecha Factura","#A11212") &_
								celdahead("Tipo Cambio","#A11212") &_
								celdahead("Proveedor","#A11212") &_
								celdahead("Incoterm","#A11212") &_
								celdahead("Factor Moneda","#A11212") &_
								celdahead("Factor","#A11212") &_
								celdahead("Proporcion","#A11212") &_
								celdahead("Adicional","#A11212") &_
								celdahead("Serie","#A11212") &_
								celdahead("Codigo de producto","#A11212") &_
								celdahead("Descripcion Item","#A11212") &_
								celdahead("Cantidad Comercial","#A11212") &_
								celdahead("Unidad de Medida Comercial","#A11212") &_
								celdahead("Cantidad Tarifa","#A11212") &_
								celdahead("Unidad de Medida Tarifa","#A11212") &_
								celdahead("Valor Moneda Extranjera","#A11212") &_
								celdahead("Valor USD","#A11212") &_
								celdahead("Categoria","#A11212") &_
								celdahead("Fraccion","#A11212") &_
								celdahead("Pais Origen","#A11212") &_
								celdahead("TLC","#A11212") &_
								celdahead("PROSEC","#A11212") &_
								celdahead("Advalorem","#A11212") &_
								celdahead("IGI","#A11212") &_
								celdahead("Tasa IVA","#A11212") &_
								celdahead("Automatico","#A11212") &_
								celdahead("Pais Vendedor","#A11212") &_
								celdahead("Tipo de Operacion","#A11212") 
														
				header = header &	"</tr>"
				'dim repCG, tieneCGD, snco,snco2, catidadph
				
			Do Until RSops.EOF
				'repCG=RSops.Fields.Item("Fecha_RecepCG_reciente").Value 
				'tieneCGD=RSops.Fields.Item("DocumentosCG").Value
				'cantidadph=Cint(RSops.Fields.Item("Cant_ComprobantesPHDigitalizados").Value)
				
					' if repCG<>"" and tieneCGD="" then
						' 'Filtro: Si la fecha de recepcion de cuenta de gastos se encuentra capturada y no se encuentra digitalizada la cuenta de gastos entonces se marcara la fila 
						' snco="#F2F5A9"
						' if cantidadph<Cint(CantPH(RSops.Fields.Item("refcia01").Value,RSops.Fields.Item("CuentasGastos").Value,mid(RSops.Fields.Item("refcia01").Value,1,3)))  then 
							' snco2="#FFFF00"
						' else
						' snco2=snco
						' end if
					' else
						 snco="#FFFFFF"
						' snco2=snco
					' end if
							datos = datos & "<tr> " &_
							celdadatos(RSops.Fields.Item("refcia01").Value,snco) &_
							celdadatos(RSops.Fields.Item("Pedimento").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Fecha de Pago").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Fecha de Entrada").Value,snco) &_ 
							 celdadatos(RSops.Fields.Item("Tipo De Cambio").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Regimen").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Clave Pedimento").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Patente").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Aduana").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Seccion").Value,snco) &_
							 celdadatos(RSops.Fields.Item("Valor Seguros").Value,snco) &_
							celdadatos(RSops.Fields.Item("Seguros").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fletes").Value,snco) &_
							celdadatos(RSops.Fields.Item("Embalajes").Value,snco)&_
							celdadatos(RSops.Fields.Item("Otros Incrementables").Value,snco)&_
							celdadatos(RSops.Fields.Item("DTA").Value,snco)&_
							celdadatos(RSops.Fields.Item("Adicional").Value,snco)&_
							celdadatos(RSops.Fields.Item("B/L").Value,snco)&_
							celdadatos(RSops.Fields.Item("Buque").Value,snco)&_
							celdadatos(RSops.Fields.Item("Descripcion Fraccion").Value,snco)&_
							celdadatos(RSops.Fields.Item("Factura").Value,snco)&_
							celdadatos(RSops.Fields.Item("Fecha Factura").Value,snco)&_
							celdadatos(RSops.Fields.Item("Tipo Cambio").Value,snco)&_
							celdadatos(RSops.Fields.Item("Proveedor").Value,snco)&_
							celdadatos(RSops.Fields.Item("Incoterm").Value,snco)&_
							celdadatos(RSops.Fields.Item("Factor Moneda").Value,snco)&_
							celdadatos(RSops.Fields.Item("Factor").Value,snco)&_
							celdadatos(RSops.Fields.Item("Proporcion").Value,snco)&_
							celdadatos(RSops.Fields.Item("Adicional").Value,snco)&_
							celdadatos(RSops.Fields.Item("Serie").Value,snco)&_
							celdadatos(RSops.Fields.Item("Codigo de Producto").Value,snco)&_
							celdadatos(RSops.Fields.Item("Descripcion Item").Value,snco)&_
							celdadatos(RSops.Fields.Item("Cantidad Comercial").Value,snco)&_
							celdadatos(RSops.Fields.Item("Unidad de Medida Comercial").Value,snco)&_
							celdadatos(RSops.Fields.Item("Cantidad Tarifa").Value,snco)&_
							celdadatos(RSops.Fields.Item("Unidad de Medida Tarifa").Value,snco)&_
							celdadatos(RSops.Fields.Item("Valor Moneda Extranjera").Value,snco)&_
							celdadatos(RSops.Fields.Item("Valor USD").Value,snco)&_
							celdadatos(RSops.Fields.Item("Categoria").Value,snco)&_
							celdadatos(RSops.Fields.Item("Fraccion").Value,snco)&_
							celdadatos(RSops.Fields.Item("Pais Origen").Value,snco)&_
							celdadatos(RSops.Fields.Item("TLC").Value,snco)&_
							celdadatos(RSops.Fields.Item("PROSEC").Value,snco)&_
							celdadatos(RSops.Fields.Item("Advalorem").Value,snco)&_
							celdadatos(RSops.Fields.Item("IGI").Value,snco)&_
							celdadatos(RSops.Fields.Item("Tasa IVA").Value,snco)&_
							celdadatos(RSops.Fields.Item("Automatico").Value,snco)&_
							celdadatos(RSops.Fields.Item("Pais Vendedor").Value,snco)&_
							celdadatos(RSops.Fields.Item("Tipo de Operacion").Value,snco)
							 datos = datos &	"</tr>"
				Rsops.MoveNext()
			Loop
	Response.Write(info & header & datos & "</table><br>")
	Response.End()
	ConnStr.Close()
	html = info & header & datos & "</table><br>"
	End If


function celdahead(texto,colorh)'Celda de encabezado de la tabla
	cell = "<td bgcolor = """&colorh&""" width=""200"" nowrap>" &_
				"<center>" &_
					"<strong>" &_
						"<font color=""#FFFFFF"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
							texto &_
						"</font>" &_
					"</strong>" &_
				"</center>" &_
			"</td>"
	celdahead = cell
end function

function celdadatos(texto,pcolor)'Celda de datos de la tabla
'On error resume next
'response.write(texto & "   ..")
texto = texto & ""
'response.end()
	If IsNull(texto) = True Or texto = "" Then
		texto = "&nbsp;"
	End If
	 dim c 
	 c=chr(34)
	cell = 	"<td align=""center""nowrap bgcolor="&c&pcolor&c&" >" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function GeneraSQL(op)
	SQL = ""
	' if strOficina <> "Todas" then
		' 'Se selecciono una oficina en especifico
		' if op="a" then 
			SQL=subSQL("IMPORTACION","i",strOficina)
			SQL= SQL & " UNION ALL "& subSQL("EXPORTACION","e",strOficina)
		' elseif op="i" then 
			' SQL=subSQL("IMPORTACION","i",strOficina)
		' elseif op="e" then 
			' SQL=subSQL("EXPORTACION","e",strOficina)
		' end if
	' elseif strOficina="Todas" then 
		' dim strOficina2
		' for ii=1 to 6
			' 'Aqui se realiza el llamado de la digitalizacion de todas las oficinas segun el tipo de operacion seleccionado
			' select case ii
				' case 1
					' strOficina2="rku"
				' case 2
					' strOficina2="sap"
				' case 3
					' strOficina2="lzr"	
				' case 5
					' strOficina2="tol"
				' case 6
					' strOficina2="ceg"
				' end select
				' if op="a" then 
					' SQL=SQL & subSQL("IMPORTACION","i",strOficina2)
					 ' SQL=SQL &" UNION ALL "& subSQL("EXPORTACION","e",strOficina2)
				' elseif op="i" then
					' SQL= SQL & subSQL("IMPORTACION",op,strOficina2)
				' elseif op="e" then
					' SQL=SQL & subSQL("EXPORTACION",op,strOficina2)
				' end if 
				' if ii < 6 then 
				 ' SQL=SQL &" UNION ALL "& chr(13) & chr(10)
				' end if
		' next
		' 'response.write(SQL)
		' 'response.end()
	' end if
	GeneraSQL = SQL
	
end function

function subSQL (operacion,movimiento,oficina)'Aqui se construye el query segun el tipo de operacion y la oficina
	SQL="SELECT ops.refcia01, "&_
		"cast(CONCAT_WS(' ', MID(ops.fecpag01, 3, 2), ops.cveadu01, ops.patent01, ops.numped01)as char ) AS 'Pedimento', "&_
		"DATE_FORMAT(ops.fecpag01, '%d/%m/%Y') AS 'Fecha de Pago', "
		if movimiento="i" then
			SQL=SQL& " DATE_FORMAT(ops.fecent01, '%d/%m/%Y') AS 'Fecha de Entrada', "
		else
			SQL=SQL& " DATE_FORMAT(ops.fecpre01, '%d/%m/%Y') AS 'Fecha de Entrada', "
		end if 
		SQL=SQL &"ops.tipcam01 AS 'Tipo De Cambio', "&_
		"ops.regime01 AS 'Regimen', "&_
		"ops.cveped01 AS 'Clave Pedimento', "&_
		"ops.patent01 AS 'Patente', "&_
		"ops.cveadu01 AS 'Aduana', "&_
		"ops.cvesec01 AS 'Seccion', "&_
		"ops.valseg01 AS 'Valor Seguros', "&_
		"ops.segros01 AS 'Seguros', "&_
		"ops.fletes01 AS 'Fletes', "&_
		"ops.embala01 AS 'Embalajes', "&_
		"ops.incble01 AS 'Otros Incrementables', "&_
		"(SELECT SUM(imp.import36) "&_
		"FROM "&oficina&"_extranet.sscont36 AS imp "&_
		"WHERE imp.refcia36 = ops.refcia01 AND imp.cveimp36 = 1 AND imp.patent36 = ops.patent01 AND imp.adusec36 = ops.adusec01 "&_
		"GROUP BY imp.refcia36) AS 'DTA', "&_
		"'' AS 'Adicional', "&_
		"IFNULL((SELECT GROUP_CONCAT(gui.numgui04 SEPARATOR ', ') "&_
		"FROM "&oficina&"_extranet.ssguia04 AS gui "&_
		"WHERE gui.refcia04 = ops.refcia01 AND gui.patent04 = ops.patent01 AND gui.adusec04 = ops.adusec01 "&_
		"GROUP BY gui.refcia04), '') AS 'B/L', "&_
		"ops.nombar01 AS 'Buque',  "&_
		"fra.d_mer102 AS 'Descripcion Fraccion', "&_
		"fac.numfac39 AS 'Factura', "&_
		"DATE_FORMAT(fac.fecfac39, '%d/%m/%Y') AS 'Fecha Factura', "&_
		"ops.tipcam01 AS 'Tipo Cambio', "&_
		"fac.nompro39 AS 'Proveedor', "&_
		"fac.terfac39 AS 'Incoterm',  "&_
		"fac.monfac39 AS 'Factor Moneda',  "&_
		"fac.facmon39 AS 'Factor',  "&_
		"'' AS 'Proporcion', "&_
		"'' AS 'Adicional',  "&_
		"'' AS 'Serie',  "&_
		"d05.cpro05 AS 'Codigo de Producto',  "&_
		"REPLACE(REPLACE(REPLACE(d05.desc05, ';', ''), '\r', ''), '\n', '') AS 'Descripcion Item',  "&_
		"d05.caco05 AS 'Cantidad Comercial',  "&_
		"d05.umco05 AS 'Unidad de Medida Comercial',  "&_
		"d05.cata05 AS 'Cantidad Tarifa',  "&_
		"d05.umta05 AS 'Unidad de Medida Tarifa', "&_
		"ROUND(d05.vafa05, 2) AS 'Valor Moneda Extranjera',  "&_
		"ROUND((d05.vafa05 * fac.facmon39), 2) AS 'Valor USD',  "&_
		"'I' AS 'Categoria',  "&_
		"d05.frac05 AS 'Fraccion',  "&_
		"fra.paiori02 AS 'Pais Origen',  "&_
		"'N' AS 'TLC', "&_
		"(SELECT GROUP_CONCAT(DISTINCT(IF(IF(par.cveide12 = 'TL', CONCAT_WS('-', par.cveide12, par.comide12), par.cveide12) = 'PS', CONCAT_WS('-', par.cveide12, par.comide12), par.cveide12)) SEPARATOR ', ') "&_
		"FROM "&oficina&"_extranet.ssipar12 AS par "&_
		"WHERE par.refcia12 = ops.refcia01 AND par.patent12 = ops.patent01 AND par.adusec12 = ops.adusec01) AS 'PROSEC', "&_
		"fra.tasadv02 AS 'Advalorem', "&_
		"IFNULL((SELECT SUM(imp.import36)  "&_
		"FROM "&oficina&"_extranet.sscont36 AS imp "&_
		"WHERE imp.refcia36 = ops.refcia01 AND imp.cveimp36 = 6 AND imp.patent36 = ops.patent01 AND imp.adusec36 = ops.adusec01 "&_
		"GROUP BY imp.refcia36), 0) AS 'IGI', "&_
		"fra.tasiva02 AS 'Tasa IVA',  "&_
		"'' AS 'Automatico',  "&_
		"fra.paiscv02 AS 'Pais Vendedor',  "
		if movimiento="i" then 
			SQL=SQL&" cast('Importacion'as char )AS 'Tipo de Operacion' "
		else 
			SQL=SQL&"cast('Exportacion' as char )AS 'Tipo de Operacion' "
		end if
		SQL=SQL& " FROM "&oficina&"_extranet.ssdag"&movimiento&"01 AS ops "&_
		"LEFT JOIN "&oficina&"_extranet.c01refer AS c01 ON c01.refe01 = ops.refcia01 "&_
		"LEFT JOIN "&oficina&"_extranet.ssfrac02 AS fra ON fra.refcia02 = ops.refcia01 AND fra.patent02 = ops.patent01 AND fra.adusec02 = ops.adusec01 "&_
		"LEFT JOIN "&oficina&"_extranet.d05artic AS d05 ON d05.refe05 = ops.refcia01 AND d05.agru05 = fra.ordfra02 AND d05.frac05 = fra.fraarn02 "&_
		"LEFT JOIN "&oficina&"_extranet.ssfact39 AS fac ON fac.refcia39 = ops.refcia01 AND fac.patent39 = ops.patent01 AND fac.adusec39 = ops.adusec01  "&_
		"AND fac.numfac39 = d05.fact05 "&_
		"WHERE ops.rfccli01 in('"&RFCliente&"') AND ops.firmae01 <> '' AND ops.firmae01 IS NOT NULL  "&_
		"AND ops.fecpag01 >= '"&DateI&"' AND ops.fecpag01 <= '"&DateF&"' AND ops.tipped01 = 1"
	
	subSQL=SQL
	
	'response.write(subSQL)
	'response.end()
end function 

%>
<HTML>
	<HEAD>
		<TITLE>::.... Anexo 24 .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>
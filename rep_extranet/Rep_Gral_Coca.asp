<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000


strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

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
	nocolumns=85
	
	For xtipo = 0 To 0
        If xtipo = 0 then
		 tablamov = "ssdagi01"
			For y = 0 To 3
				If y = 0 Then
					strOficina = "rku"
					query = GeneraSQL
	                query = query & " union all "
				else 
					if y = 1 then
						strOficina = "sap"
						query = query & GeneraSQL
	                    query = query & " union all "
					else
						if y = 2 then
							strOficina = "dai"
							query = query & GeneraSQL
	                        query = query & " union all "
						 else
						   if y = 3 then
							strOficina = "tol"
							query = query & GeneraSQL
					       End If
						End If
					End If
				End If
			Next
		else
		 
		end if
    Next
	
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	Set RSops = CreateObject("ADODB.RecordSet")

	Set RSops = ConnStr.Execute(query)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
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
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>" &_
												" IMPORTACIONES" &_
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
				"</table>"
		
		header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">" &_
							    celdahead("REF. ZEGO") &_
								celdahead("ITEM") &_
								celdahead("VENDOR") &_
								celdahead("PRODUCTO") &_
								celdahead("PROVEEDOR") &_
								celdahead("NOMBRE COMERCIAL") &_
								celdahead("DGR") &_
								celdahead("ADUANA") &_
								celdahead("TLC") &_
								celdahead("PAÍS DE ORIGEN") &_
								celdahead("PAÍS DE PROCED") &_
								celdahead("ORDEN DE COMPRA") &_
								celdahead("LINEA DE ORDEN DE COMPRA") &_
								celdahead("NÚMERO DE FACTURA") &_
								celdahead("FECHA DE FACTURA") &_
								celdahead("PESO NETO KG") &_
								celdahead("VALOR UNITARIO") &_
								celdahead("GASTOS EXTRANJERO OTROS USD") &_
								celdahead("VALOR FACTURA EUR/GBP/JYN") &_
								celdahead("FACTOR DE CONVERSION") &_
								celdahead("VALOR FACTURA USD") &_
								celdahead("BULTOS CNTRS") &_
								celdahead("MONEDA FACTURADO ORIGEN") &_
								celdahead("FLETE MX/ EUR/ JYN/ GBP") &_
								celdahead("FLETE INTERNAL USD") &_
								celdahead("NUMERO DE FACTURA TRANSPORTISTA") &_
								celdahead("LINEA TRANSPORTE INTERNACIONAL") &_
								celdahead("No. DE GUIA AWB / BL USA") &_
								celdahead("LUGAR DE ENTREGA MEX") &_
								celdahead("FECHA DE ETD") &_
								celdahead("FECHA ARRIBO A FRONTERA") &_
								celdahead("FECHA DE DESPACHO") &_
								celdahead("FECHA ARRIBO A PLANTA MEX") &_
								celdahead("FECHA DE ENTREGA DE CONTENEDOR VACÍO") &_
								celdahead("FLETE NACIONAL MÉXICO CPS") &_
								celdahead("LINEA DE TRANSPORTE MEXICANA") &_
								celdahead("No. DE GUIA AWB / BL MEX") &_
								celdahead("No. DE PEDIMENTO") &_
								celdahead("PEDIMENTO TIPO DE CAMBIO") &_
								celdahead("PESO BRUTO DECLARADO EN PEDIMENTO") &_
								celdahead("INCOTERM") &_
								celdahead("USD INCREMENTAB OTROS") &_
								celdahead("USD VALOR EN DÓLARES") &_
								celdahead("VALOR ADUANA") &_
								celdahead("VALOR COMERCIAL") &_
								celdahead("INCREMENTAB  FLETES") &_
								celdahead("INCREMENT. OTROS") &_
								celdahead("FRACCIÓN ARANCELARIA") &_
								celdahead("%TASA DTA") &_
								celdahead("DTA") &_
								celdahead("TASA %IGI") &_
								celdahead("IGI") &_
								celdahead("PREV") &_
								celdahead("%IVA") &_
								celdahead("IVA") &_
								celdahead("TOTAL DE IMPUESTOS") &_
								celdahead("NÚMERO FACTURA CTA GASTOS") &_
								celdahead("DESPACHO MANIOBRAS") &_
								celdahead("DEMORAS") &_
								celdahead("SERVICIOS COMPLEMENTARIOS") &_
								celdahead("HONORARIOS") &_
								celdahead("MONTO DE FLETE PAGADO MNX AEREO") &_
								celdahead("CUENTA AMERICANA USD LAREDO") &_
								celdahead("CUENTA AMERICANA MXN") &_
								celdahead("SOLICITUD DE ANTICIPO A LA A.A") &_
								celdahead("INGRESO DE ANTICIPO") &_
								celdahead("VALOR ANTICIPO") &_
								celdahead("MONTO TOTAL CTA GASTOS") &_
								celdahead("SALDOS") &_
								celdahead("FECHA ENTREGA CUENTA DE GASTOS") &_
								celdahead("MPORTACIÓN EXPORTACIÓN") &_
								celdahead("MES") &_
								celdahead("DESADUANAMIENTO") &_
								celdahead("NÚMERO DE UN IATA") &_
								celdahead("NÚMERO DE UN QN") &_
								celdahead("COMENTARIOS") &_
								celdahead("FLASH POINT") &_
								celdahead("TIPO DE CAMBIO ESTÁNDAR") &_
								celdahead("FLETE INTERNAL / KG o PZA") &_
								celdahead("FLETE NACIONAL / KG O PZA") &_
								celdahead("GASTOS DE IMPORTACION / KG O PZA") &_
								celdahead("IMPUESTOS SIN PREF / KG O PZA") &_
								celdahead("IMPUESTOS REAL / KG O PZA") &_
								celdahead("TOTAL SIN PREF / KG O PZA") &_
								celdahead("TOTAL CON PREF / KG O PZA") 
		header = header &	"</tr>"
		datos = ""
		Referencia = ""
		contador=5
		Do Until RSops.EOF
			contador = contador + 1
			Referencia = RSops.Fields.Item("refcia01")
			datos = datos &	"<tr>" &_
			celdadatos(RSops.Fields.Item("refcia01").Value) &_
			celdadatos(RSops.Fields.Item("Item").Value) &_
			celdadatos(RSops.Fields.Item("Vendor").Value) &_
			celdadatos(RSops.Fields.Item("Producto").Value) &_
			celdadatos(RSops.Fields.Item("Proveedor").Value) &_
			celdadatos(RSops.Fields.Item("NombreComercial").Value) &_
			celdadatos(RSops.Fields.Item("DGR").Value) &_
			celdadatos(RSops.Fields.Item("Aduana").Value) &_
			celdadatos(RSops.Fields.Item("TLC").Value) &_
			celdadatos(RSops.Fields.Item("PaisOrigen").Value) &_
			celdadatos(RSops.Fields.Item("PaisProcedencia").Value) &_
			celdadatos(RSops.Fields.Item("OrdenCompra").Value) &_
			celdadatos(RSops.Fields.Item("LineaOrdenCompra").Value) &_
			celdadatos(RSops.Fields.Item("NumeroFactura").Value) &_
			celdadatos(RSops.Fields.Item("FechaFactura").Value) &_
			celdadatos(RSops.Fields.Item("PesoNeto").Value) &_
			celdadatos(RSops.Fields.Item("Valor Unitario").Value) &_
			celdadatos(RSops.Fields.Item("GastosExtrangOtrDls").Value) &_
			celdadatos(RSops.Fields.Item("VALOR FACTURA EUR/GBP/JYN").Value) &_
			celdadatos(RSops.Fields.Item("Factor de conversion").Value) &_
			celdadatos(RSops.Fields.Item("Valor Dls").Value) &_
			celdadatos(RSops.Fields.Item("Bultos/CNTRS").Value) &_
			celdadatos(RSops.Fields.Item("MonedaFacturadoOrigen").Value) &_
			celdadatos(RSops.Fields.Item("FleteMX/EUR/JYN").Value) &_
			celdadatos(RSops.Fields.Item("FleteInternacionalUSD").Value) &_
			celdadatos(RSops.Fields.Item("NoFacturaTransportista").Value) &_
			celdadatos(RSops.Fields.Item("LineaTransporteInternacional").Value) &_
			celdadatos(RSops.Fields.Item("GuiaAWB/BL USA").Value) &_
			celdadatos(RSops.Fields.Item("LugarEntregaAMex").Value) &_
			celdadatos(RSops.Fields.Item("FechaETD").Value) &_
			celdadatos(RSops.Fields.Item("FecheArriboFrontera").Value) &_
			celdadatos(RSops.Fields.Item("FechaDespacho").Value) &_
			celdadatos(RSops.Fields.Item("FechaArriboAPlantaMex").Value) &_
			celdadatos(RSops.Fields.Item("FechaentregaContenedorVacio").Value) &_
			celdadatos(RSops.Fields.Item("FleteNacionalMexicoCPS").Value) &_
			celdadatos(RSops.Fields.Item("LineaTransporteMexicana").Value) &_
			celdadatos(RSops.Fields.Item("GuiaAWB/BL MEX").Value) &_
			celdadatos(RSops.Fields.Item("NumeroPedimento").Value) &_
			celdadatos(RSops.Fields.Item("TipoCambioPedimento").Value) &_
			celdadatos(RSops.Fields.Item("PesoBrutoDecenPed").Value) &_
			celdadatos(RSops.Fields.Item("INCOTERM").Value) &_
			celdadatos(RSops.Fields.Item("IncremOtrsDls").Value) &_
			celdadatos(RSops.Fields.Item("ValorenDolares").Value) &_
			celdadatos(RSops.Fields.Item("ValorAduana").Value) &_
			celdadatos(RSops.Fields.Item("ValorComercial").Value) &_
			celdadatos(RSops.Fields.Item("IncrementablesFletes").Value) &_
			celdadatos(RSops.Fields.Item("IncrementablesOtros").Value) &_
			celdadatos(RSops.Fields.Item("FraccionAranc").Value) &_
			celdadatos(RSops.Fields.Item("% DTA").Value & "%") &_
			celdadatos(RSops.Fields.Item("DTA").Value) &_
			celdadatos(RSops.Fields.Item("%IGI").Value & "%") &_
			celdadatos(RSops.Fields.Item("IGI").Value) &_
			celdadatos(RSops.Fields.Item("PRV").Value) &_
			celdadatos(RSops.Fields.Item("%IVA").Value & "%") &_
			celdadatos(RSops.Fields.Item("IVA").Value) &_
			celdadatos(RSops.Fields.Item("TotalImpuestos").Value) &_
			celdadatos(RSops.Fields.Item("cgas31").Value) &_
			celdadatos(RSops.Fields.Item("Maniobras").Value) &_
			celdadatos(RSops.Fields.Item("Demoras").Value) &_
			celdadatos(RSops.Fields.Item("ServiciosComplementarios").Value) &_
			celdadatos(RSops.Fields.Item("HONORARIOS").Value) &_
			celdadatos(RSops.Fields.Item("MontoFletePagMNXAereo").Value) &_
			celdadatos(RSops.Fields.Item("CuentaAmericanaUDSLaredo").Value) &_
			celdadatos(RSops.Fields.Item("CuentaAmericanaMXN").Value) &_
			celdadatos(RSops.Fields.Item("SolicitudDeAnticipoALaAA").Value) &_
			celdadatos(RSops.Fields.Item("IngresodeAnticipo").Value) &_
			celdadatos(RSops.Fields.Item("ValorAnticipo").Value) &_
			celdadatos(RSops.Fields.Item("MontoTotalCtaGastos").Value) &_
			celdadatos(RSops.Fields.Item("Saldos").Value) &_
			celdadatos(RSops.Fields.Item("FechaEntregaCtaGastos").Value) &_
			celdadatos(RSops.Fields.Item("ImportacionExportacion").Value) &_
			celdadatos(RSops.Fields.Item("Mes").Value) &_
			celdadatos(RSops.Fields.Item("Desaduanamiento").Value) &_
			celdadatos(RSops.Fields.Item("NumeroDeUnIATA").Value) &_
			celdadatos(RSops.Fields.Item("NumeroDeQN").Value) &_
			celdadatos(RSops.Fields.Item("Comentarios").Value) &_
			celdadatos(RSops.Fields.Item("FlashPoint").Value) &_
			celdadatos(RSops.Fields.Item("TipoCambioEstandar").Value) &_
			celdadatos("=Y"&cstr(contador) & "/ P" & cstr(contador)) &_
			celdadatos("=(AI"&cstr(contador) & "/ BZ" & cstr(contador) & ")/ P" & cstr(contador)) &_
			celdadatos("=((BF"&cstr(contador) & "+ BG" & cstr(contador) & "+ BH" & cstr(contador) & "+ BI" & cstr(contador) & ")/ BZ" & cstr(contador) & ")/ P" & cstr(contador)) &_
			celdadatos("=((AR"&cstr(contador) & "/ BZ" & cstr(contador) & ") * 100%) /P" & cstr(contador)) &_
			celdadatos("=((AX"&cstr(contador) & "+ AZ" & cstr(contador) & "+ BA" & cstr(contador) & ")/ BZ" & cstr(contador) & ")/ P" & cstr(contador)) &_
			celdadatos("=CA"&cstr(contador) & "+ CB" & cstr(contador) & "+ CC" & cstr(contador) & "+ CD" & cstr(contador)) &_
			celdadatos("=CA"&cstr(contador) & "+ CB" & cstr(contador) & "+ CC" & cstr(contador) & "+ CE" & cstr(contador))
						
			datos = datos &	"</tr>"
			Rsops.MoveNext()
		Loop
	  
	 'prom = ""
	 'prom = Promedios
	
 	 html = info & header & datos & "</table><br>"
	
	
	End If
end if

function GeneraSQL
	SQL = ""
	condicion = filtro
	SQL = 	"SELECT i.refcia01," &_
			"'' Item, " &_
			"'' Vendor, " &_
			"fr.d_mer102 Producto, " &_
			"i.nompro01 Proveedor, " &_
			"replace(replace(replace(ar.desc05,'\r',''),'\n',''),char(13),'') NombreComercial, " &_
			"if(oem.pgro01 = 'T','Y',if(oem.pgro01 = 'F','N','')) DGR, " &_
			"case i.adusec01 " &_
			"when 430 then 'VERACRUZ' " &_
			"when 160 then 'MANZANILLO' " &_
			"when 510 then 'CARDENAS' " &_
			"when 650 then 'TOLUCA' " &_
			"when 200 then 'MEXICO' " &_
			"when 202 then 'MEXICO' " &_
			"when 470 then 'AICM' " &_
			"when 472 then 'AICM' " &_
			"when 810 then 'ALTAMIRA' " &_
			"when 380 then 'TAMPICO' " &_
			"ELSE 'OTRO' " &_
			"END AS 'Aduana', " &_
			"( select ifnull(group_concat(ipar2.cveide12,ipar2.comide12),'') " &_
			"  from " & strOficina & "_extranet.ssipar12 ipar2  " &_
			"  where ipar2.refcia12 =i.refcia01 and ipar2.patent12 = i.patent01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 ='TL' ) TLC," &_
			"i.cvepod01 PaisOrigen, " &_
			"( select ptoem.Cvepai01 " &_
			"  from " & strOficina & "_extranet.c01ptoemb ptoem " &_
			"  where ptoem.Cvepto01 = r.cveptoemb and ptoem.Nompto01 = r.ptoemb01) PaisProcedencia," &_
			"ar.pedi05 OrdenCompra, " &_
			"'' LineaOrdenCompra, " &_
			"f.numfac39 NumeroFactura, " &_
			"f.fecfac39 FechaFactura," &_
			"i.cmucom01 PesoNeto, " &_
			"(ar.vafa05 / ar.caco05) 'Valor Unitario', " &_
			"'' GastosExtrangOtrDls, " &_
			"f.valmex39 'VALOR FACTURA EUR/GBP/JYN', " &_
			"f.facmon39  'Factor de conversion', " &_
			"f.valdls39 'Valor Dls', " &_
			"(select concat(sum(x1.cant01),' (',group_concat(distinct  x1.clas01),')') from " & strOficina & "_extranet.d01conte as x1 where x1.refe01 =i.refcia01) 'Bultos/CNTRS', " &_
			"f.monfac39 MonedaFacturadoOrigen, " &_
			"'' AS 'FleteMX/EUR/JYN', " &_
			"(i.fletes01 / i.tipcam01) 'FleteInternacionalUSD', " &_
			"'' NoFacturaTransportista, " &_
			"'' LineaTransporteInternacional, " &_
			"ifnull(gui2.numgui04,'') 'GuiaAWB/BL USA', " &_
			"ifnull(replace(replace(replace(oem.obse01,'\r',''),'\n',''),char(13),''),'') LugarEntregaAMex, " &_
			"'' FechaETD, " &_
			"r.feta01 FecheArriboFrontera, " &_
			"r.fdsp01 FechaDespacho, " &_
			"'' FechaArriboAPlantaMex, " &_
			"( select	ct.fcarta01 " &_
			"  from " & strOficina & "_extranet.d01conte ct " &_
			"  where		ct.refe01 = i.refcia01 " &_
			"  group by	ct.fcarta01 ) FechaentregaContenedorVacio, " &_
			"'' FleteNacionalMexicoCPS, " &_
			"IFNULL(( select group_concat(trans.nom02) " &_
			"  from " & strOficina & "_extranet.d01conte d01c " &_
			"   left join " & strOficina & "_extranet.e01oemb oem on d01c.peri01 = oem.peri01 and d01c.nemb01 = oem.nemb01  " &_
			"    left join " & strOficina & "_extranet.c56trans trans on oem.ctra01 = trans.cve02 " &_
			"  where d01c.refe01 = i.refcia01),'') LineaTransporteMexicana, " &_
			"gui1.numgui04 as 'GuiaAWB/BL MEX', " &_
			"concat(i.patent01 ,' ',i.numped01) NumeroPedimento, " &_
			"i.tipcam01 TipoCambioPedimento, " &_
			"i.pesobr01 PesoBrutoDecenPed, " &_
			"f.terfac39 INCOTERM, " &_
			"(i.incble01 / i.tipcam01) IncremOtrsDls, " &_
			"i.valdol01 ValorenDolares, " &_
			"(i.valdol01 * i.tipcam01) ValorAduana, " &_
			"(i.valmer01 * i.tipcam01) ValorComercial, " &_
			"i.fletes01 IncrementablesFletes, " &_
			"i.incble01 IncrementablesOtros, " &_
			"fr.fraarn02 FraccionAranc, " &_
			"i.tsadta01 '% DTA', " &_
			"ifnull(cf1.import36,0) DTA, " &_
			"fr.tasadv02 AS '%IGI', " &_
			"ifnull(cf6.import36,0) IGI, " &_
			"cf15.import36 PRV, " &_
			"fr.tasiva02 AS '%IVA', " &_
			"cf3.import36 'IVA', " &_
			"(ifnull(cf1.import36,0)  + ifnull(cf3.import36,0) + ifnull(cf6.import36,0) + ifnull(cf15.import36,0) ) TotalImpuestos, " &_
			"cta.cgas31, " &_
			Maniobras(strOficina) &_
			Demoras(strOficina) &_
			ServiciosComplementarios(strOficina) &_
			"  cta.chon31 HONORARIOS,	 " &_
			"'' MontoFletePagMNXAereo, " &_
			"'' CuentaAmericanaUDSLaredo, " &_
			"'' CuentaAmericanaMXN, " &_
			"r.fcot01 SolicitudDeAnticipoALaAA, " &_
			"mvant.fech11 IngresodeAnticipo, " &_
			"ifnull(mvant.mont11,0) ValorAnticipo, " &_
			MontoTotalCtaGastos(strOficina) &_
			Saldos(strOficina) &_
			"cta.fech31 FechaEntregaCtaGastos, " &_
			"'IMPORTACION' AS ImportacionExportacion, " &_
			"case month(i.fecpag01) " &_
			"when 1 then 'ENERO' " &_
			"when 2 then 'FEBRERO' " &_
			"when 3 then 'MARZO' " &_
			"when 4 then 'ABRIL' " &_
			"when 5 then 'MAYO' " &_
			"when 6 then 'JUNIO' " &_
			"when 7 then 'JULIO' " &_
			"when 8 then 'AGOSTO' " &_
			"when 9 then 'SEPTIEMBRE' " &_
			"when 10 then 'OCTUBRE' " &_
			"when 11 then 'NOVIEMBRE' " &_
			"when 12 then 'DICIEMBRE' " &_
			"END  Mes, " &_
			"(SELECT IF(SUM(IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1,0))=0,'LIBRE','RECONOCIMIENTO') " &_
			"  FROM trackingbahia.bit_soia AS bs " &_
			"    LEFT JOIN trackingbahia.cat_situaciones AS cs ON bs.cvesit01 = cs.cvesit01 " &_
			"     LEFT JOIN trackingbahia.cat_det_situaciones AS cds ON cds.detsit01 = bs.detsit01 " &_
			"  where bs.frmsaai01 = i.refcia01 " &_
			") Desaduanamiento, " &_
			"'' NumeroDeUnIATA, " &_
			"'' NumeroDeQN, " &_
			"'' Comentarios, " &_
			"'' FlashPoint, " &_
			"case f.monfac39 " &_
			"when 'USD' then '12.4497' " &_
			"when 'EUR' then '17.4904' " &_
			"when 'STG' then '20.0069' " &_
			"else '' " &_
			"end as 'TipoCambioEstandar' " &_
			"from " & strOficina & "_extranet." & tablamov & " as i " &_
			" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
			"  left join " & strOficina & "_extranet.c01refer as r on r.refe01 = i.refcia01 " &_
			"   LEFT join " & strOficina & "_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01  " &_
			"    inner join " & strOficina & "_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " &_
			"     left join " & strOficina & "_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01   " &_
			"      left join " & strOficina & "_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01  " &_
			"        left join " & strOficina & "_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05    " &_
			"         left join " & strOficina & "_extranet.ssguia04 as gui2 on gui2.refcia04 = i.refcia01 and gui2.patent04 = i.patent01 and gui2.adusec04 = i.adusec01 and gui2.idngui04 =2 " &_
			"          left join " & strOficina & "_extranet.ssguia04 as gui1 on gui1.refcia04 = i.refcia01 and gui1.patent04 = i.patent01 and gui1.adusec04 = i.adusec01 and gui1.idngui04 =1 " &_
			"           left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1' " &_
			"            left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3' " &_
			"             left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6' " &_
			"              left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15' " &_
			"               left join " & strOficina & "_extranet.d11movim as mvant on mvant.refe11 = i.refcia01 and mvant.conc11 = 'ANT' and mvant.cgas11 = ctar.cgas31  " &_
			"                left join " & strOficina & "_extranet.d01conte as d01c on d01c.refe01 = i.refcia01 " &_
			"                 left join " & strOficina & "_extranet.e01oemb oem on d01c.peri01 = oem.peri01 and d01c.nemb01 = oem.nemb01  " &_
			"where i.firmae01 is not null and i.firmae01 <> '' " &_
			"and i.fecpag01 >= '" & DateI & "' and i.fecpag01 <= '" & DateF & "' " & condicion &_
			" group by i.refcia01, cta.cgas31,f.numfac39,ar.item05 "
			

	GeneraSQL = SQL
end function

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
		texto = "-"
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
		condicion = " and cc.rfccli18 = '" & Vrfc & "' "
	'else
	'	if Vclave <> "Todos" Then
	'		condicion = "AND i.cvecli01 = " & Vclave & " "
	'	Else
	'		permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
	'		condicion = permi
	'		condicion = "AND " & condicion
	'		if condicion = "AND cvecli01=0 " Then
	'			condicion = ""
	'		end if
	'	End If
	end if
	'condicion = " and cc.rfccli18 = 'CCE520101TC7' "
	filtro = condicion
end function

Function Maniobras (oficina)
cadena = ""
if oficina="rku" then
 cadena =   "ifnull(( select sum((if(epage.deha21 ='C',-1,1) ) *dpage.mont21/ (1+(epage.piva21/100)  )) " &_
			"  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage  " &_
			"  where dpage.refe21 = i.refcia01 " &_
			"  and dpage.cgas21 = ctar.cgas31 " &_
			"  and epage.foli21 = dpage.foli21 " &_
			"  and year(epage.fech21) = year(dpage.fech21) " &_
			"  and epage.esta21 <> 'S' " &_
			"  and epage.esta21 <> 'C' " &_
			"  and epage.tmov21 = dpage.tmov21 " &_
			"  and epage.conc21  in (2)),0) Maniobras, "
else
    if oficina="sap" then
		 cadena =   "ifnull(( select sum((if(epage.deha21 ='C',-1,1) ) *dpage.mont21/ (1+(epage.piva21/100)  )) " &_
					"  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage  " &_
					"  where dpage.refe21 = i.refcia01 " &_
					"  and dpage.cgas21 = ctar.cgas31 " &_
					"  and epage.foli21 = dpage.foli21 " &_
					"  and year(epage.fech21) = year(dpage.fech21) " &_
					"  and epage.esta21 <> 'S' " &_
					"  and epage.esta21 <> 'C' " &_
					"  and epage.tmov21 = dpage.tmov21 " &_
					"  and epage.conc21  in (2,70,111,144,182)),0) Maniobras, "
	else
		if oficina="dai" then
			 cadena =   "ifnull(( select sum((if(epage.deha21 ='C',-1,1) ) *dpage.mont21/ (1+(epage.piva21/100)  )) " &_
						"  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage  " &_
						"  where dpage.refe21 = i.refcia01 " &_
						"  and dpage.cgas21 = ctar.cgas31 " &_
						"  and epage.foli21 = dpage.foli21 " &_
						"  and year(epage.fech21) = year(dpage.fech21) " &_
						"  and epage.esta21 <> 'S' " &_
						"  and epage.esta21 <> 'C' " &_
						"  and epage.tmov21 = dpage.tmov21 " &_
						"  and epage.conc21  in (2,11,22,82,127,232)),0) Maniobras, "
		else
			if oficina="tol" then
			 cadena =   "ifnull(( select sum((if(epage.deha21 ='C',-1,1) ) *dpage.mont21/ (1+(epage.piva21/100)  )) " &_
						"  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage  " &_
						"  where dpage.refe21 = i.refcia01 " &_
						"  and dpage.cgas21 = ctar.cgas31 " &_
						"  and epage.foli21 = dpage.foli21 " &_
						"  and year(epage.fech21) = year(dpage.fech21) " &_
						"  and epage.esta21 <> 'S' " &_
						"  and epage.esta21 <> 'C' " &_
						"  and epage.tmov21 = dpage.tmov21 " &_
						"  and epage.conc21  in (127,2)),0) Maniobras, "
			end if
		end if
	end if
end if
Maniobras = cadena
end Function

Function Demoras (oficina)
cadena = ""
if oficina="rku" then
 cadena =   "ifnull(( select sum(dpage.mont21) " &_
			"  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage " &_
			"  where dpage.refe21 = i.refcia01 " &_
			"  and dpage.cgas21 = ctar.cgas31 " &_
			"  and epage.foli21 = dpage.foli21 " &_
			"  and year(epage.fech21) = year(dpage.fech21) " &_
			"  and epage.esta21 <> 'S' " &_
			"  and epage.esta21 <> 'C' " &_
			"  and epage.tmov21 = dpage.tmov21 " &_
			"  and epage.conc21 = 11),0) Demoras, "
else
    if oficina="sap" then
		 cadena =   "ifnull(( select sum(dpage.mont21) " &_
					"  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage " &_
					"  where dpage.refe21 = i.refcia01 " &_
					"  and dpage.cgas21 = ctar.cgas31 " &_
					"  and epage.foli21 = dpage.foli21 " &_
					"  and year(epage.fech21) = year(dpage.fech21) " &_
					"  and epage.esta21 <> 'S' " &_
					"  and epage.esta21 <> 'C' " &_
					"  and epage.tmov21 = dpage.tmov21 " &_
					"  and epage.conc21 = 6),0) Demoras, "
	else
		if oficina="dai" then
			 cadena =   "0 Maniobras, "
		else
			if oficina="tol" then
			 cadena =   "0 Maniobras, "
			end if
		end if
	end if
end if
Demoras = cadena
end Function

Function ServiciosComplementarios (oficina)
cadena = ""
if oficina="rku" then
 cadena =   "  cta.caho31 ServiciosComplementarios, "
else
    if oficina="sap" then
		 cadena =   "  cta.caho31 ServiciosComplementarios, "
	else
		if oficina="dai" then
			 cadena =   "ifnull(( select sum((if(epage.deha21 ='C',-1,1) ) *dpage.mont21/ (1+(epage.piva21/100)  )) " &_
					    "  from " & oficina & "_extranet.d21paghe dpage, " & oficina & "_extranet.e21paghe epage " &_
						"where dpage.refe21 = i.refcia01 " &_
						"and dpage.cgas21 = ctar.cgas31 " &_
						"and epage.foli21 = dpage.foli21 " &_
						"and year(epage.fech21) = year(dpage.fech21) " &_
						"and epage.esta21 <> 'S' " &_
						"and epage.esta21 <> 'C' " &_
						"and epage.tmov21 = dpage.tmov21 " &_
						"and epage.conc21  in (140)),0) + (cta.csce31 + cta.caho31) ServiciosComplementarios, "
		else
			if oficina="tol" then
			 cadena =   "(cta.csce31 + cta.caho31 ) ServiciosComplementarios, "
			end if
		end if
	end if
end if
ServiciosComplementarios = cadena
end Function

Function MontoTotalCtaGastos (oficina)
cadena = ""
if oficina="rku" then
 cadena =   "  (((cta.chon31 + cta.caho31) * (1 + (cta.piva31/100))) + (cta.suph31)) MontoTotalCtaGastos, "
else
    if oficina="sap" then
		 cadena =   "  (((cta.chon31 + cta.caho31) * (1 + (cta.piva31/100))) + (cta.suph31)) MontoTotalCtaGastos, "
	else
		if oficina="dai" then
			 cadena =   " (((cta.chon31 + (cta.csce31 + cta.caho31 )) * (1 + (cta.piva31/100))) + (cta.suph31)) MontoTotalCtaGastos, "
		else
			if oficina="tol" then
			 cadena =   " (((cta.chon31 + (cta.csce31 + cta.caho31 )) * (1 + (cta.piva31/100))) + (cta.suph31)) MontoTotalCtaGastos, "
			end if
		end if
	end if
end if
MontoTotalCtaGastos = cadena
end Function

Function Saldos (oficina)
cadena = ""
if oficina="rku" then
 cadena =   "  ((((cta.chon31 + cta.caho31) * (1 + (cta.piva31/100))) + (cta.suph31)) - ifnull(mvant.mont11,0)) Saldos, "
else
    if oficina="sap" then
		 cadena =   "  ((((cta.chon31 + cta.caho31) * (1 + (cta.piva31/100))) + (cta.suph31)) - ifnull(mvant.mont11,0)) Saldos, "
	else
		if oficina="dai" then
			 cadena =   " ((((cta.chon31 + (cta.csce31 + cta.caho31 )) * (1 + (cta.piva31/100))) + (cta.suph31)) - ifnull(mvant.mont11,0)) Saldos, "
		else
			if oficina="tol" then
			 cadena =   " ((((cta.chon31 + (cta.csce31 + cta.caho31 )) * (1 + (cta.piva31/100))) + (cta.suph31)) - ifnull(mvant.mont11,0)) Saldos, "
			end if
		end if
	end if
end if
Saldos = cadena
end Function

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
			"FROM " & strOficina & "_extranet.d01conte AS d01 " &_
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
				"FROM " & strOficina & "_extranet.log_actualiza " &_
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
		<TITLE>::.... REPORTE GENERAL DE THE COCA-COLA EXPORT CORPORATION SUC. EN MEXICO.... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
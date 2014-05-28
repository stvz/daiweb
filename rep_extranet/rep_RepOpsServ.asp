<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% 
strTipoUsuario = request.Form("TipoUser") '004 ejecutivo cliente
strPermisos = Request.Form("Permisos")

permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

if not permi = "" then
	permi = "  AND (" & permi & ") "
end if
AplicaFiltro = False
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	blnAplicaFiltro = true
end if
if blnAplicaFiltro then
	permi = " AND cvecli01 =" & strFiltroCliente & " "
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	permi = ""
end if
%>
<html>
	<head>
		<meta http-equiv=Content-Type content="text/html; charset=utf-8">
		<meta name=ProgId content=Excel.Sheet>
		<meta name=Generator content="Microsoft Excel 11">
		<title>REPORTE OPERACIONES Y FLETES</title>
	</head>
	<body>
	<% 
	

	
	
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
		multiofi = Request.Form("multi")
		filtro = ""
		refe = ""
		nofrac = 0
		cont = 0
		
		If strTipoUsuario = "004" Then 
			Vckcve = 1
		End If
		if multiofi = "S" And vckcve <> 0 Then
			if strTipoUsuario <> "004" Then
				Response.Write("No se puede realizar multioficina y seleccionando Clave de cliente, por favor elija RFC")
			Else
				Response.Write("No se puede realizar multioficina ya que las claves de cliente varian dependiendo" &_
								"de la oficina en la que se encuentra, solicite permisos para hacer consultas por RFC")
			End If
			Response.End()
		End If
		
		filtro = "WHERE "
		
		
		If IsDate(fi) = True Then
			DiaI = cstr(datepart("d",fi))
			Mesi = cstr(datepart("m",fi))
			AnioI = cstr(datepart("yyyy",fi))
			DateI = Anioi & "/" & Mesi & "/" & Diai
			filtro = filtro & "i.fecpag01 >= '" & DateI &"' "
		End If
		
		IF IsDate(ff) = True then
			DiaF = cstr(datepart("d",ff))
			MesF = cstr(datepart("m",ff))
			AnioF = cstr(datepart("yyyy",ff))
			DateF = AnioF & "/" & MesF & "/" & DiaF
			filtro = filtro & "AND i.fecpag01 <= '" & DateF & "' "
		End If
		
		If Vckcve = 0 Then
			filtro = filtro & "AND i.rfccli01 like '" & Vrfc & "' "
		Else
			filtro = filtro & permi
		End If
		
		filtro = filtro & "AND i.firmae01 <> '' AND i.cveped01 <> 'R1' "
		' Response.Write(filtro)
		' Response.End()
		query = GeneraSQL(filtro)
		' Response.Write(query)
		' Response.End
		tabla = ""
		Set StrConn = Server.CreateObject("ADODB.Connection")
		StrConn.Open = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = Server.CreateObject("ADODB.Recordset")
		Set RSops = StrConn.Execute(query)
		If RSops.BOF = True And RSops.EOF = True Then
			Response.Write("No hay datos que mostrar")
			Response.End()
		Else
			tabla = 	"<Row>" &_
							GeneraCelda("String", "Referencia") &_
							GeneraCelda("String", "Compañia") &_
							GeneraCelda("String", "Agencia Aduanal") &_
							GeneraCelda("String", "Factura AA") &_
							GeneraCelda("String", "Pedimento") &_
							GeneraCelda("String", "Tipo Pedimento") &_
							GeneraCelda("String", "DLLS") &_
							GeneraCelda("String", "T.C.") &_
							GeneraCelda("String", "M/N") &_
							GeneraCelda("String", "Valor Declarado en Aduana") &_
							GeneraCelda("String", "Gastos Incrementales") &_
							GeneraCelda("String", "IGI") &_
							GeneraCelda("String", "DTA") &_
							GeneraCelda("String", "Base IVA") &_
							GeneraCelda("String", "Tasa IVA") &_
							GeneraCelda("String", "IVA Importacion") &_
							GeneraCelda("String", "PRV") &_
							GeneraCelda("String", "Total del pedimento") &_
							GeneraCelda("String", "Descripcion") &_
							GeneraCelda("String", "Cantidad Importada") &_
							GeneraCelda("String", "Medida") &_
							GeneraCelda("String", "Nombre Pais Origen") &_
							GeneraCelda("String", "Nombre Pais Destino") &_
							GeneraCelda("String", "Clave De Pedimento") &_
							GeneraCelda("String", "Nombre Del Proveedor Extranjera") &_
							GeneraCelda("String", "Factura Proveedor Extranjero") &_
							GeneraCelda("String", "TAX ID") &_
							GeneraCelda("String", "Numero De Cuenta De Gastos") &_
							GeneraCelda("String", "Nombre Del Agente Aduanal") &_
							GeneraCelda("String", "CUENTA 0605101000 16% P4") &_
							GeneraCelda("String", "CUENTA 0605101000 11% P3") &_
							GeneraCelda("String", "CUENTA 0605101000 0% P6") &_
							GeneraCelda("String", "16%") &_
							GeneraCelda("String", "11%") &_
							GeneraCelda("String", "RETENCION FLETES 4%") &_
							GeneraCelda("String", "TOTAL CUENTA DE GASTOS") &_
							GeneraCelda("String", "ANTICIPO") &_
							GeneraCelda("String", "SALDO DE CUENTA DE GASTOS") &_
							GeneraCelda("String", "Numero Factura") &_
							GeneraCelda("String", "Nombre Fletero") &_
							GeneraCelda("String", "RFC Del Fletero") &_
							GeneraCelda("String", "Tasa IVA") &_
							GeneraCelda("String", "Base De Retencion") &_
							GeneraCelda("String", "Otros Gastos") &_
							GeneraCelda("String", "IVA") &_
							GeneraCelda("String", "10% HONORARIOS") &_
							GeneraCelda("String", "10.66% HONORARIOS") &_
							GeneraCelda("String", "4% Fletes") &_
							GeneraCelda("String", "Total") &_
						"</Row>"
			RSops.MoveFirst()
			Do Until RSops.EOF = True
				tabla = tabla & GeneraFila
				RSops.MoveNext()
			Loop
			DIM XML : XML = "<?xml version='1.0'?>" &_
							"<?mso-application progid='Excel.Sheet'?>" &_
							"<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' " &_
							"xmlns:o='urn:schemas-microsoft-com:office:office' " &_
							"xmlns:x='urn:schemas-microsoft-com:office:excel' " &_
							"xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet' " &_
							"xmlns:html='http://www.w3.org/TR/REC-html40'>" &_
								"<DocumentProperties xmlns='urn:schemas-microsoft-com:office:office'>" &_
									"<Author>Rogelio Gonzalez</Author>" &_
									"<LastAuthor>Rogelio Gonzalez</LastAuthor>" &_
									"<Created>2010-12-15T22:30:45Z</Created>" &_
									"<Company>Grupo ZEGO</Company>" &_
									"<Version>11.5606</Version>" &_
								"</DocumentProperties>" &_
								"<ExcelWorkbook xmlns='urn:schemas-microsoft-com:office:excel'>" &_
									"<WindowHeight>12270</WindowHeight>" &_
									"<WindowWidth>15195</WindowWidth>" &_
									"<WindowTopX>480</WindowTopX>" &_
									"<WindowTopY>45</WindowTopY>" &_
									"<ProtectStructure>False</ProtectStructure>" &_
									"<ProtectWindows>False</ProtectWindows>" &_
								"</ExcelWorkbook>" &_
								"<Styles>" &_
									"<Style ss:ID='Default' ss:Name='Normal'>" &_
										"<Alignment ss:Vertical='Bottom'/>" &_
										"<Borders/>" &_
										"<Font/>" &_
										"<Interior/>" &_
										"<NumberFormat/>" &_
										"<Protection/>" &_
									"</Style>" &_
								"</Styles>" &_
								"<Worksheet ss:Name='Hoja1'>" &_
									"<Table ss:ExpandedColumnCount='50' ss:ExpandedRowCount='2500' x:FullColumns='1' " &_
									"x:FullRows='1' ss:DefaultColumnWidth='60'>" &_
										tabla &_
									"</Table>" &_
									"<WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" &_
										"<PageSetup>" &_
											"<Header x:Margin='0'/>" &_
											"<Footer x:Margin='0'/>" &_
											"<PageMargins x:Bottom='0.984251969' x:Left='0.78740157499999996' " &_
											"x:Right='0.78740157499999996' x:Top='0.984251969'/>" &_
										"</PageSetup>" &_
										"<Selected/>" &_
										"<Panes>" &_
											"<Pane>" &_
												"<Number>3</Number>" &_
												"<ActiveCol>2</ActiveCol>" &_
											"</Pane>" &_
										"</Panes>" &_
										"<ProtectObjects>False</ProtectObjects>" &_
										"<ProtectScenarios>False</ProtectScenarios>" &_
									"</WorksheetOptions>" &_
								"</Worksheet>" &_
							"</Workbook>"
			With Response
				.Clear
				.ContentType = "excel/ms-excel"
				.AddHeader "Content-Disposition","attachment; filename =""Rep Operaciones Fletes.xml"""
				.AddHeader "Content-Length", Len(XML)
				.Write XML
				.Flush
				.End
			End With
		End If
	End If
	
	Function GeneraSQL(filtrado)
		SQL =	""
		
		if multiofi = "S" Then
			conta = 1
		Else
			conta = 6
		End If
		For conta = conta To 6
			If multiofi = "S" Then
				Select case conta
					case 1
						strOficina = "rku"
					Case 2
						strOficina = "dai"
					Case 3
						strOficina = "sap"
					Case 4
						strOficina = "lzr"
					Case 5
						strOficina = "ceg"
					Case 6
						strOficina = "tol"
				End Select
			End If
			SQL = 	SQL & "SELECT i.refcia01 AS 'referencia', " &_
					"if(i.rfccli01 LIKE 'CCE520101TC7', 'Coca-Cola', 'JDV Marcko') AS 'compania', " &_
					"(SELECT c00.raso00 FROM " & strOficina & "_extranet.c00trage AS c00 WHERE LENGTH(c00.raso00)>15) AS 'AgenciaA', " &_
					"(SELECT GROUP_CONCAT(d31.cgas31) FROM " & strOficina & "_extranet.d31refer AS d31 " &_
						"INNER JOIN " & strOficina & "_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 " &_
						"WHERE d31.refe31 = i.refcia01 AND e31.esta31 <> 'C' GROUP BY d31.refe31) AS 'facturaAA', " &_
					"CONCAT(i.patent01, i.numped01) AS 'pedimento', " &_
					"'IMP' AS 'tipoped', " &_
					"FORMAT((SELECT SUM(fac.valdls39) FROM " & strOficina & "_extranet.ssfact39 as fac " &_
						"WHERE fac.refcia39=i.refcia01 AND i.adusec01 = fac.adusec39 AND fac.patent39 = i.patent01 " &_
						"GROUP BY fac.refcia39), 2) AS 'vfactura', " &_
					"i.tipcam01 AS 'tipocam', " &_
					"IFNULL(i.fletes01, 0) AS 'incrementables', " &_
					"IFNULL((SELECT SUM(adv.import36) FROM " & strOficina & "_extranet.sscont36 AS adv WHERE adv.refcia36 = i.refcia01 AND adv.cveimp36 = 6 ), 0) AS 'IGI', " &_
					"IFNULL((SELECT SUM(dta.import36) FROM " & strOficina & "_extranet.sscont36 AS dta WHERE dta.refcia36 = i.refcia01 AND dta.cveimp36 = 1 ), 0) AS 'DTA', " &_
					"(SELECT COUNT(DISTINCT(CONCAT_WS('-', fr.fraarn02, fr.ordfra02))) " &_
						"FROM " & strOficina & "_extranet.ssfrac02 AS fr WHERE fr.refcia02 = i.refcia01 AND fr.adusec02 = i.adusec01 AND fr.patent02 = i.patent01) AS 'nfrac', " &_
					"(fr.i_adv102 + fr.i_adv202 + fr.i_adv302) AS 'ADV', " &_
					"(fr.tasiva02 / 100) AS 'tiva', " &_
					"(fr.i_iva102 + fr.i_iva202 + fr.i_iva302) AS 'IVA', " &_ 
					"IFNULL((SELECT SUM(prv.import36) FROM " & strOficina & "_extranet.sscont36 AS prv WHERE prv.refcia36 = i.refcia01 AND prv.cveimp36 = 15 GROUP BY prv.refcia36), 0) AS 'prv', " &_
					"REPLACE(REPLACE(fr.d_mer102, '\r', ''), '\n', '') AS 'despartida', " &_
					"fr.cancom02 AS 'cantidad', " &_
					"(SELECT med.descri31 FROM " & strOficina & "_extranet.ssumed31 AS med WHERE med.clavem31 = fr.u_medc02) AS 'unimed', " &_
					"fr.paiscv02 AS 'porigen', " &_
					"fr.paiori02 AS 'pdestino', " &_
					"i.cveped01 AS 'cveped', " &_
					"(SELECT GROUP_CONCAT(DISTINCT(fac.nompro39) SEPARATOR ', ') FROM " & strOficina & "_extranet.ssfact39 AS fac " &_
						"WHERE fac.refcia39 = i.refcia01 AND fac.adusec39 = i.adusec01 AND fac.patent39 = i.patent01) AS 'proveedor', " &_
					"(SELECT GROUP_CONCAT(fac.numfac39 SEPARATOR ', ') FROM " & strOficina & "_extranet.ssfact39 AS fac " &_
						"WHERE fac.refcia39 = i.refcia01 AND fac.adusec39 = i.adusec01 AND fac.patent39 = i.patent01) AS 'facturas', " &_
					"(SELECT GROUP_CONCAT(DISTINCT(fac.idfisc39) SEPARATOR ', ') FROM " & strOficina & "_extranet.ssfact39 AS fac " &_
						"WHERE fac.refcia39 = i.refcia01 AND fac.adusec39 = i.adusec01 AND fac.patent39 = i.patent01) AS 'idfiscal', " &_
					"(SELECT group_concat(distinct con.agente32) FROM " & strOficina & "_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AgenteA', " &_
					"IFNULL((SELECT (SUM(dp.mont21 * IF(ep.deha21 = 'C', -1,1))/((ep.piva21 / 100) + 1)) " &_
						"FROM " & strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " & strOficina & "_extranet.e21paghe AS ep ON dp.foli21 = ep.foli21 AND YEAR(dp.fech21) = YEAR(ep.fech21)  " &_
							"AND ep.esta21 <> 'S' AND ep.esta21 <> 'C'  AND ep.tmov21 =dp.tmov21 AND ep.piva21 = 16 " &_
						"LEFT JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 NOT LIKE '%DEPO%E%GARANT%' GROUP BY dp.refe21), 0) AS 'pagos16', " &_
					"IFNULL((SELECT (SUM(dp.mont21 * IF(ep.deha21 = 'C', -1,1))/((ep.piva21 / 100) + 1)) " &_
						"FROM " & strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " & strOficina & "_extranet.e21paghe AS ep ON dp.foli21 = ep.foli21 AND YEAR(dp.fech21) = YEAR(ep.fech21) " &_
							"AND ep.esta21 <> 'S' AND ep.esta21 <> 'C'  AND ep.tmov21 =dp.tmov21 AND ep.piva21 = 11 " &_
						"LEFT JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 NOT LIKE '%DEPO%E%GARANT%' GROUP BY dp.refe21), 0) AS 'pagos11', " &_
					"IFNULL((SELECT (SUM(dp.mont21 * IF(ep.deha21 = 'C', -1,1))/((ep.piva21 / 100) + 1)) " &_
						"FROM " & strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " & strOficina & "_extranet.e21paghe AS ep ON dp.foli21 = ep.foli21 AND YEAR(dp.fech21) = YEAR(ep.fech21) " &_
							"AND ep.esta21 <> 'S' AND ep.esta21 <> 'C'  AND ep.tmov21 =dp.tmov21 AND ep.piva21 = 0 " &_
						"LEFT JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 NOT LIKE '%DEPO%E%GARANT%' GROUP BY dp.refe21), 0) AS 'pagos0', " &_
					"(SELECT SUM(IFNULL(e31.tota31, 0)) FROM " &strOficina & "_extranet.d31refer AS d31 " &_
						"LEFT JOIN " &strOficina & "_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
						"WHERE d31.refe31 = i.refcia01) AS 'totacg', " &_
					"(SELECT SUM(IF(d11.conc11 = 'ANT', 1, -1) * d11.mont11) FROM " &strOficina & "_extranet.d11movim AS d11 " &_
						"WHERE d11.refe11 = i.refcia01 AND d11.conc11 IN ('CAN', 'ANT') GROUP BY d11.refe11) AS 'anticipo', " &_
					"((SELECT SUM(IFNULL(e31.tota31, 0)) FROM " &strOficina & "_extranet.d31refer AS d31 " &_
						"LEFT JOIN " &strOficina & "_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 <> 'C' " &_
						"WHERE d31.refe31 = i.refcia01) - " &_
						"(SELECT SUM(IF(d11.conc11 = 'ANT', 1, -1) * d11.mont11) FROM " &strOficina & "_extranet.d11movim AS d11 " &_
						"WHERE d11.refe11 = i.refcia01 AND d11.conc11 IN ('CAN', 'ANT') GROUP BY d11.refe11)) as 'saldo', " &_
					"(SELECT GROUP_CONCAT(DISTINCT(CONCAT(dp.facpro21)) SEPARATOR ' / ') FROM " &strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " &strOficina & "_extranet.e21paghe AS ep ON ep.foli21 = dp.foli21 AND YEAR(ep.fech21) = YEAR(dp.fech21) " &_
						"AND ep.esta21 <> 'S' AND ep.esta21 <> 'C' AND ep.tmov21 = dp.tmov21 " &_
						"INNER JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 LIKE '%FLETE%TERRESTR%') AS 'factfletes', " &_
					"(SELECT GROUP_CONCAT(DISTINCT(ben.nomb20) SEPARATOR ' / ') FROM " &strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " &strOficina & "_extranet.e21paghe AS ep ON ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) " &_
						"AND ep.esta21 <> 'S' AND ep.esta21 <> 'C' AND ep.tmov21 = dp.tmov21 " &_
						"INNER JOIN " &strOficina & "_extranet.c20benef AS ben ON ben.clav20 = ep.bene21 AND ben.aplic20 = 'F' " &_
						"INNER JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 LIKE '%FLETE%TERRESTR%' GROUP BY dp.refe21) AS 'beneficiario', " &_
					"(SELECT GROUP_CONCAT(DISTINCT(ben.rfc20) SEPARATOR ' / ') FROM " &strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " &strOficina & "_extranet.e21paghe AS ep ON ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) " &_
						"AND ep.esta21 <> 'S' AND ep.esta21 <> 'C' AND ep.tmov21 = dp.tmov21 " &_
						"INNER JOIN " &strOficina & "_extranet.c20benef AS ben ON ben.clav20 = ep.bene21 AND ben.aplic20 = 'F' " &_
						"INNER JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 LIKE '%FLETE%TERRESTR%' GROUP BY dp.refe21) AS 'rfcbenef', " &_
					"((SELECT GROUP_CONCAT(DISTINCT(ep.piva21) SEPARATOR ' / ') FROM " &strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " &strOficina & "_extranet.e21paghe AS ep ON ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) " &_
						"AND ep.esta21 <> 'S' AND ep.tmov21 = dp.tmov21 " &_
						"INNER JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 LIKE '%FLETE%TERRESTR%' GROUP BY dp.refe21) / 100) AS 'tivafletes', " &_
					"(SELECT SUM(mfle21 * (IF(ep.deha21 = 'C',-1,1))) FROM " &strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " &strOficina & "_extranet.e21paghe AS ep ON ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) " &_
						"AND ep.esta21 <> 'S' AND ep.tmov21 = dp.tmov21 " &_
						"INNER JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 LIKE '%FLETE%TERRESTR%' GROUP BY dp.refe21) AS 'baseflete', "&_
					"(SELECT SUM((((dp.mont21 + (dp.mfle21*0.04))/1.16) - dp.mfle21) * (IF(ep.deha21 = 'C',-1,1))) FROM " &strOficina & "_extranet.d21paghe AS dp " &_
						"INNER JOIN " &strOficina & "_extranet.e21paghe AS ep ON ep.foli21 = dp.foli21 and YEAR(ep.fech21) = YEAR(dp.fech21) " &_
						"AND ep.esta21 <> 'S' AND ep.tmov21 = dp.tmov21 " &_
						"INNER JOIN " &strOficina & "_extranet.c21paghe AS cp ON cp.clav21 = ep.conc21 " &_
						"WHERE dp.refe21 = i.refcia01 AND cp.desc21 LIKE '%FLETE%TERRESTR%' GROUP BY dp.refe21) AS 'otros' " &_
					"FROM " & strOficina & "_extranet.ssdagi01 AS i " &_
					"LEFT JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON fr.refcia02 = i.refcia01 AND fr.adusec02 = i.adusec01 AND fr.patent02 = i.patent01 " &_
					filtrado
					if conta <> 6 then
						SQL = SQL & "UNION ALL "
					End If
		Next
		SQL = SQL & "ORDER BY referencia "
		' Response.Write(SQL)
		' Response.End()
		GeneraSQL = SQL
	End Function
	
	Function GeneraCelda(tipo, valor)
		cell = ""
		cell =	"<Cell><Data ss:Type='" & tipo & "'>" & valor & "</Data></Cell>"
		' Response.Write(cell)
		' Response.End()
		GeneraCelda = cell
	End Function
	
	Function GeneraFormula(formula, tipo)
		celda =	""
		celda =	"<Cell ss:Formula='" & formula & "'><Data ss:Type='" & tipo & "'></Data></Cell>"
		GeneraFormula = celda
	End Function
	
	Function GeneraFila
		Fila = ""
		If refe <> RSops("referencia") then
			refe = RSops("referencia")
			nofrac = RSops("nfrac")
			If nofrac = 1 Then
				Fila =	"<Row>" &_
							GeneraCelda("String", RSops("referencia")) &_
							GeneraCelda("String", RSops("compania")) &_
							GeneraCelda("String", RSops("AgenciaA")) &_
							GeneraCelda("String", RSops("facturaAA")) &_
							GeneraCelda("String", RSops("pedimento")) &_
							GeneraCelda("String", RSops("tipoped")) &_
							GeneraCelda("Number", RSops("vfactura")) &_
							GeneraCelda("Number", RSops("tipocam")) &_
							GeneraFormula("=RC[-2]*RC[-1]", "Number") &_
							GeneraFormula("=+RC[-1]", "Number") &_
							GeneraCelda("Number", RSops("incrementables")) &_
							GeneraCelda("Number", RSops("IGI")) &_
							GeneraCelda("Number", RSops("DTA")) &_
							GeneraFormula("=SUM(RC[-4]:RC[-1])", "Number") &_
							GeneraCelda("Number", RSops("tiva")) &_
							GeneraFormula("=(RC[-2]*RC[-1])", "Number") &_
							GeneraCelda("Number", RSops("PRV")) &_
							GeneraFormula("=RC[-6]+RC[-5]+RC[-2]+RC[-1]", "Number") &_
							GeneraCelda("String", RSops("despartida")) &_
							GeneraCelda("Number", RSops("cantidad")) &_
							GeneraCelda("String", RSops("unimed")) &_
							GeneraCelda("String", RSops("porigen")) &_
							GeneraCelda("String", RSops("pdestino")) &_
							GeneraCelda("String", RSops("cveped")) &_
							GeneraCelda("String", RSops("proveedor")) &_
							GeneraCelda("String", RSops("facturas")) &_
							GeneraCelda("String", RSops("idfiscal")) &_
							GeneraFormula("=RC[-24]", "String") &_
							GeneraCelda("String", RSops("AgenteA")) &_
							GeneraCelda("Number", RSops("pagos16")) &_
							GeneraCelda("Number", RSops("pagos11")) &_
							GeneraCelda("Number", RSops("pagos0")) &_
							GeneraFormula("=RC[-3]*0.16", "Number") &_
							GeneraFormula("=RC[-3]*0.11", "Number") &_
							GeneraFormula("=RC[13]", "Number") &_
							GeneraCelda("Number", RSops("totacg")) &_
							GeneraCelda("Number", RSops("anticipo")) &_
							GeneraCelda("Number", RSops("saldo")) &_
							GeneraCelda("String", RSops("factfletes")) &_
							GeneraCelda("String", RSops("beneficiario")) &_
							GeneraCelda("String", RSops("rfcbenef")) &_
							GeneraCelda("Number", 0.16) &_
							GeneraCelda("Number", RSops("baseflete")) &_
							GeneraCelda("Number", RSops("otros")) &_
							GeneraFormula("=(SUM(RC[-2]:RC[-1])*RC[-3])", "Number") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraFormula("=(-RC[-5]*0.04)", "Number") &_
							GeneraFormula("=SUM(RC[-6]:RC[-1])", "Number") &_
						"</Row>"
			Else
				Fila =	"<Row>" &_
							GeneraCelda("String", refe) &_
							GeneraCelda("String", RSops("compania")) &_
							GeneraCelda("String", RSops("AgenciaA")) &_
							GeneraCelda("String", RSops("facturaAA")) &_
							GeneraCelda("String", RSops("pedimento")) &_
							GeneraCelda("String", RSops("tipoped")) &_
							GeneraCelda("Number", RSops("vfactura")) &_
							GeneraCelda("Number", RSops("tipocam")) &_
							GeneraFormula("=RC[-2]*RC[-1]", "Number") &_
							GeneraFormula("=+RC[-1]", "Number") &_
							GeneraCelda("Number", RSops("incrementables")) &_
							GeneraFormula("=SUM(R[1]C:R[" & nofrac & "]C)", "Number") &_
							GeneraCelda("Number", RSops("DTA")) &_
							GeneraFormula("=SUM(RC[-4]:RC[-1])", "Number") &_
							GeneraCelda("String", "") &_
							GeneraFormula("=SUM(R[1]C:R[" & nofrac & "]C)", "Number") &_
							GeneraCelda("Number", RSops("PRV")) &_
							GeneraFormula("=RC[-6]+RC[-5]+RC[-2]+RC[-1]", "Number") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", RSops("cveped")) &_
							GeneraCelda("String", RSops("proveedor")) &_
							GeneraCelda("String", RSops("facturas")) &_
							GeneraCelda("String", RSops("idfiscal")) &_
							GeneraFormula("=RC[-24]", "String") &_
							GeneraCelda("String", RSops("AgenteA")) &_
							GeneraCelda("Number", RSops("pagos16")) &_
							GeneraCelda("Number", RSops("pagos11")) &_
							GeneraCelda("Number", RSops("pagos0")) &_
							GeneraFormula("=RC[-3]*0.16", "Number") &_
							GeneraFormula("=RC[-3]*0.11", "Number") &_
							GeneraFormula("=RC[13]", "Number") &_
							GeneraCelda("Number", RSops("totacg")) &_
							GeneraCelda("Number", RSops("anticipo")) &_
							GeneraCelda("Number", RSops("saldo")) &_
							GeneraCelda("String", RSops("factfletes")) &_
							GeneraCelda("String", RSops("beneficiario")) &_
							GeneraCelda("String", RSops("rfcbenef")) &_
							GeneraCelda("Number", 0.16) &_
							GeneraCelda("Number", RSops("baseflete")) &_
							GeneraCelda("Number", RSops("otros")) &_
							GeneraFormula("=(SUM(RC[-2]:RC[-1])*RC[-3])", "Number") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraFormula("=(-RC[-5]*0.04)", "Number") &_
							GeneraFormula("=SUM(RC[-6]:RC[-1])", "Number") &_
						"</Row>" &_
						"<Row>" &_
							GeneraCelda("String", refe) &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("Number", RSops("ADV")) &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("Number", RSops("tiva")) &_
							GeneraCelda("Number", RSops("IVA")) &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", RSops("despartida")) &_
							GeneraCelda("Number", RSops("cantidad")) &_
							GeneraCelda("String", RSops("unimed")) &_
							GeneraCelda("String", RSops("porigen")) &_
							GeneraCelda("String", RSops("pdestino")) &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
							GeneraCelda("String", "") &_
						"</Row>"
			End If
		Else
			Fila =	"<Row>" &_
						GeneraCelda("String", refe) &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("Number", RSops("ADV")) &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("Number", RSops("tiva")) &_
						GeneraCelda("Number", RSops("IVA")) &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", RSops("despartida")) &_
						GeneraCelda("Number", RSops("cantidad")) &_
						GeneraCelda("String", RSops("unimed")) &_
						GeneraCelda("String", RSops("porigen")) &_
						GeneraCelda("String", RSops("pdestino")) &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
						GeneraCelda("String", "") &_
					"</Row>"
		End If
		GeneraFila = Fila
	End Function
	%>
	</body>
</html>
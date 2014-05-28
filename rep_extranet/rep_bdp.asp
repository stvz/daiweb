<html>
	<head>
		<meta http-equiv=Content-Type content="text/html; charset=utf-8">
		<meta name=ProgId content=Excel.Sheet>
		<meta name=Generator content="Microsoft Excel 11">
		<%
		
		dim oficina,cvesoficina,validacion,codigo, cvecli, adu
		cvecli= ""
		codigo = ""
		oficina=""
		cvesoficina=""
		validacion=""

		'oficina=Request.QueryString("ofi")
		'tipope=Request.QueryString("tipope")
		'det=Request.QueryString("det")
		'mes=Request.QueryString("mes")

		strTipoUsuario = Session("GTipoUsuario")
		fechaini = trim(request.Form("txtDateIni"))
		fechafin = trim(request.Form("txtDateFin"))
		adu = trim(request.Form("Aduana"))
		Select Case adu
			case "VER"
				Oficina="rku"
			case "MEX"
				Oficina="dai"
			case "MAN"
				Oficina="sap"
			case "GUA"
				Oficina="rku"
			case "TAM"
				Oficina="ceg"
			case "LAR"
				Oficina="LAR"
			case "LZR"
				Oficina="lzr"
			case "TOL"
				Oficina="tol"
		End Select

		cvecli = trim(Request.Form("txtCliente"))
		if cvecli = "Todos" then
			cvecli = "4284"
		End If
		strTipoOperaciones = request.Form("rbnTipoDate")
		 
		 
		if not fechaini="" and not fechafin="" then

			dim finicio,ffinal
			tmpDiaIni = cstr(datepart("d",fechaini))
			tmpMesIni = cstr(datepart("m",fechaini))
			tmpAnioIni = cstr(datepart("yyyy",fechaini))
			finicio = tmpAnioIni & "-" &tmpMesIni & "-"& tmpDiaIni
			tmpDiaFin = cstr(datepart("d",fechafin))
			tmpMesFin = cstr(datepart("m",fechafin))
			tmpAnioFin = cstr(datepart("yyyy",fechafin))
			ffinal = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

			dim strHTML 
			strHTML = ""

		Server.ScriptTimeOut=100000
		%>
<title>Layout CARGILL</title>
</head>
<body>
		<% Dim XML : XML = "<?xml version='1.0'?>" & _
							"<?mso-application progid='Excel.Sheet'?>" & _
							"<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' " & _
							" xmlns:o='urn:schemas-microsoft-com:office:office' "& _
							" xmlns:x='urn:schemas-microsoft-com:office:excel' " & _
							"xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet' " & _
							" xmlns:html='http://www.w3.org/TR/REC-html40'>" & _
							" <DocumentProperties xmlns='urn:schemas-microsoft-com:office:office'>" & _
							"    <Author>Rogelio Gonzlez</Author>" & _
							"    <LastAuthor>Rogelio Gonzlez</LastAuthor>" & _
							"    <Created>" & ISODate & "</Created>" & _
							"    <Company>Grupo Zego</Company>" & _
						    "    <Version>12.00</Version>" & _
						    "</DocumentProperties>" & _
						    "<OfficeDocumentSettings xmlns='urn:schemas-microsoft-com:office:office'>" & _
						    "    <DownloadComponents/>" & _
							"</OfficeDocumentSettings>" & _
						   "<ExcelWorkbook xmlns='urn:schemas-microsoft-com:office:excel'>" & _
						   "    <WindowHeight>10425</WindowHeight>" & _
						   "    <WindowWidth>18015</WindowWidth>" & _
						   "    <WindowTopX>240</WindowTopX>" & _
						   "    <WindowTopY>60</WindowTopY>" & _
						   "    <ProtectStructure>False</ProtectStructure>" & _
						   "    <ProtectWindows>False</ProtectWindows>" & _
						   "</ExcelWorkbook>" & _
						   "<Styles>" & _
						   "    <Style ss:ID='Default' ss:Name='Normal'>" & _
						   "        <Alignment ss:Vertical='Bottom'/>" & _
						   "        <Borders/>" & _
						   "        <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='11' ss:Color='#000000'/>" & _
						   "        <Interior/>" & _
						   "        <NumberFormat/>" & _
						   "        <Protection/>" & _			   
						   "    </Style>" & _
						   "    <Style ss:ID='ds1'>" & _
						   "        <Alignment ss:Vertical='Bottom'/>" & _
						   "        <Borders/>" & _
						   "        <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='11' ss:Color='#000000'/>" & _
						   "        <NumberFormat/>" & _
						   "    </Style>" & _
						  "    <Style ss:ID='dm1'>" & _
						   "        <Alignment ss:Vertical='Bottom'/>" & _
						   "        <Borders/>" & _
						   "        <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='11' ss:Color='#000000'/>" & _
						   "<NumberFormat " &_
						   "ss:Format='_-&quot;$&quot;* #,##0.00_-;\-&quot;$&quot;* #,##0.00_-;_-&quot;$&quot;* &quot;-&quot;??_-;_-@_-'/>" &_
						   "    </Style>" & _
						   "<Style ss:ID='ds2'> " & _
						   " <Interior ss:Color='#99FFFF' ss:Pattern='Solid'></Interior> " & _
						   "  <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='11' ss:Color='#000000'/>" & _
						   "  </Style> " & _
						   "<Style ss:ID='dm2'> " & _
						   " <Interior ss:Color='#99FFFF' ss:Pattern='Solid'></Interior> " & _
						   "  <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='11' ss:Color='#000000'/>" & _
						   "<NumberFormat " &_
						   "ss:Format='_-&quot;$&quot;* #,##0.00_-;\-&quot;$&quot;* #,##0.00_-;_-&quot;$&quot;* &quot;-&quot;??_-;_-@_-'/>" &_
						   "  </Style> " & _
						    " <Style ss:ID='s21'> " & _
						   " <Interior ss:Color='#330099' ss:Pattern='Solid'></Interior> " & _
						   "  <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='12' ss:Color='#FFFFFF'/>" & _
						   "  </Style> " & _
						   " <Style ss:ID='ft1'> " & _
						   " <Interior ss:Color='#330099' ss:Pattern='Solid'></Interior> " & _
						   "  <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='12' ss:Color='#FFFFFF'/>" & _
						   "<NumberFormat " &_
						   "ss:Format='_-&quot;$&quot;* #,##0.00_-;\-&quot;$&quot;* #,##0.00_-;_-&quot;$&quot;* &quot;-&quot;??_-;_-@_-'/>" &_
						   "  </Style> " & _
						   "</Styles>" & _
						   "<Worksheet ss:Name='Layout Cargill'>" & _
									" <Table x:FullColumns='1' " & _
									" x:FullRows='1' ss:DefaultRowHeight='15'>" & _
									"<Column ss:Width='102.75'/>" &_
									"<Column ss:Width='98.25'/>" &_
									"<Column ss:Width='57'/>" &_
									"<Column ss:Width='60.75'/>" &_
									"<Column ss:Width='79.5'/>" &_
									"<Column ss:Width='87.75'/>" &_
									"<Column ss:Width='99.75'/>" &_
									"<Column ss:Width='116.25'/>" &_
									"<Column ss:Width='42'/>" &_
									"<Column ss:Width='66'/>" &_
									"<Column ss:Width='59.25'/>" &_
									"<Column ss:Width='127.5'/>" &_
									"<Column ss:Width='126'/>" &_
									"<Column ss:Width='120.75'/>" &_
									"<Column ss:Width='54.75'/>" &_
									"<Column ss:Width='63.75'/>" &_
									"<Column ss:Width='52.5'/>" &_
									"<Column ss:Width='59.25'/>" &_
									"<Column ss:Width='68.25'/>" &_
									"<Column ss:Width='57'/>" &_
									"<Column ss:Width='94.5'/>" &_
									"<Column ss:Width='70.5'/>" &_
									"<Column ss:Width='81'/>" &_
									"<Column ss:Width='34.5'/>" &_
									"<Column ss:Width='66.75'/>" &_
									"<Column ss:Width='34.5'/>" &_
									"<Column ss:Width='111.75'/>" &_
									"<Column ss:Width='34.5'/>" &_
									"<Column ss:Width='125.25'/>" &_
									"<Column ss:Width='42.75'/>" &_
									"<Column ss:Width='153.75'/>" &_
									"<Column ss:Width='47.25'/>" &_
									"<Column ss:Width='72.75'/>" &_
									"<Column ss:Width='147'/>" &_
									"<Column ss:Width='75'/>" &_
									"<Column ss:Width='61.5'/>" &_
									genera_registros("ENC","IMPO",finicio,ffinal) & _
									" </Table>" & _
					   "<WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
					   "    <PageSetup>" & _
					   "        <Header x:Margin='0.3'/>" & _
					   "        <Footer x:Margin='0.3'/>" & _
					   "        <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
					   "    </PageSetup>" & _
					   "    <Selected/>" & _
					   "    <Panes>" & _
					   "        <Pane>" & _
					   "           <Number>3</Number>" & _
					   "           <ActiveRow>0</ActiveRow>" & _
					   "           <ActiveCol>0</ActiveCol>" & _
					   "       </Pane>" & _
					   "    </Panes>" & _
					   "    <ProtectObjects>False</ProtectObjects>" & _
					   "    <ProtectScenarios>False</ProtectScenarios>" & _
					   "</WorksheetOptions>" & _
					   "</Worksheet>" & _
					   "</Workbook>"
					   
			'XML = Replace(XML, "<", "(")
			'XML = Replace(XML, ">", ")")
			
			'XML= XML & "</body> </HTML>"
			'response.Write(XML)
			'response.End()
			 With Response
			   .Charset = "UTF-8"
			   .Clear
			   .ContentType = "excel/ms-excel"
			   .AddHeader "Content-Disposition","attachment; filename=LayoutCargill.xls"
			   .AddHeader "Content-Length", Len(XML)
			   .Write XML
			   .Flush
			   .End
			 End With
		%>
</body>
</html>

<% end if

function genera_registros(det,tipope,finicio,ffinal)
	dim c,nparte, banderita, refe, color,acumvmext,acumvmn, acumincfle, acumincseg, acumincotr, acumdecfle
	dim acumdecseg, acumdecotr, acumvadu, acumpre, acumigi, acumdta, acumeci, acumimp, acumcuo, acumiva, acumtadu
	
	
	nparte=""
	color="2"
	codigo = ""
	refe=""
	banderita = 0
	acumvmext = 0
	acumvmn = 0
	acumincfle = 0
	acumincseg = 0
	acumincotr = 0
	acumdecfle = 0
	acumdecseg = 0
	acumdecotr = 0
	acumvadu = 0
	acumpre = 0
	acumigi = 0
	acumdta = 0
	acumeci = 0
	acumimp = 0
	acumcuo = 0
	acumiva = 0
	acumtadu = 0
	c=chr(34)
	 
	codigo=codigo & "<Row>"
	genera_html "e","Referencia","center","'String'"
	genera_html "e","Patente Pedimento","center","'String'"
	genera_html "e","Pedimento","center","'String'"
	genera_html "e","Fecha Pago","center","'String'"
	genera_html "e","Tipo Operacion","center","'String'"
	genera_html "e","Clave Pedimento","center","'String'"
	genera_html "e","Pedimento Original","center","'String'"
	genera_html "e","Fecha de Rectificacion","center","'String'"
	genera_html "e","Aduana","center","'String'"
	genera_html "e","Tipo Cambio","center","'String'"
	genera_html "e","INCOTERM","center","'String'"
	genera_html "e","Valor Moneda Extranjera","center","'String'"
	genera_html "e","Descripcion","right","'String'"
	genera_html "e","Valor Moneda Nacional","center","'String'"
	genera_html "e","Inc. Fletes","center","'String'"
	genera_html "e","Inc. Seguros","center","'String'"
	genera_html "e","Inc. Otros","center","'String'"
	genera_html "e","Dec. Fletes" ,"center","'String'"
	genera_html "e","Dec. Seguros" ,"center","'String'"
	genera_html "e","Dec. Otros" ,"center","'String'"
	genera_html "e","Valor Aduana" ,"center","'String'"
	genera_html "e","Prevalidacion","center","'String'"
	genera_html "e","Advalorem / IGI","center","'String'"
	genera_html "e","DTA","center","'String'"
	genera_html "e","ECI","center","'String'"
	genera_html "e","Otros","center","'String'"
	genera_html "e","Cuota Compensatoria","center","'String'"
	genera_html "e","IVA","center","'String'"
	genera_html "e","Total Pagado en Aduana","center","'String'"
	genera_html "e","Patente","center","'String'"
	genera_html "e","Agente Aduanal","center","'String'"
	genera_html "e","Factura","center","'String'"
	genera_html "e","Fecha Factura","center","'String'"
	genera_html "e","Proveedor","center","'String'"
	genera_html "e","Pais Vendedor","center","'String'"
	genera_html "e","Pais Origen","center","'String'"
	codigo=codigo & "</Row>"


	sqlAct= "SELECT " &_
			"i.refcia01 as 'Referencia', " &_
			"CONCAT_WS('-',i.patent01,i.numped01) as 'PatPed', " &_
			"i.numped01 as 'Pedi', " &_
			"DATE_FORMAT(i.fecpag01,'%d/%m/%Y') as 'FecPag', " &_
			"'1' as 'TipoOp', " &_
			"i.cveped01 as 'CvePed', " &_
			"IF (rec.pedorg06 IS NOT NULL OR rec.pedorg06 <> '',CONCAT(rec.agente06,'-',rec.pedorg06),'----') as 'pedorg', " &_
			"IF (rec.f_pagr06 IS NOT NULL OR rec.f_pagr06 <> '',DATE_FORMAT(rec.f_pagr06,'%d/%m/%Y'),'----') as 'FechaR1', " &_
			"CONCAT(i.cveadu01,'/',i.cvesec01) as 'aduana', " &_
			"IF(i.tipcam01 IS NOT NULL and i.tipcam01 <> 0,FORMAT(i.tipcam01,5),0) as 'TipoCam', " &_
			"fac.terfac39 as 'incoterm', " &_
			"IF(i.valmer01 IS NOT NULL and i.valmer01 <> '',FORMAT(i.valmer01,2),0)as 'valormext', " &_
			"IF(c.dmer01 IS NOT NULL and c.dmer01 <> '',REPLACE(REPLACE(c.dmer01,'\r',''),'\n',''),'----') as 'descrip', " &_
			"FORMAT(IF(i.valmer01 IS NULL,0,i.valmer01)*IF(i.tipcam01 IS NULL,1,i.tipcam01)*IF(i.factmo01 IS NULL,i.factmo01,1),0) as 'valormn', " &_
			"IF(i.fletes01 IS NULL OR i.fletes01 = '',0,FORMAT(i.fletes01,2)) as 'IncFletes', " &_
			"IF(i.segros01 IS NULL OR i.segros01 = '',0,FORMAT(i.segros01,2)) as 'IncSeguros', " &_
			"IF(i.incble01 IS NULL OR i.incble01 = '',0,FORMAT(i.incble01,2)) as 'IncOtros', " &_
			"IF(c.decfle01 IS NULL OR c.decfle01 = '',0,FORMAT(c.decfle01,2)) as 'DecFletes', " &_
			"IF(c.decsgr01 IS NULL OR c.decsgr01 = '',0,FORMAT(c.decsgr01,2)) as 'DecSeguros', " &_
			"IF(c.decotr01 IS NULL OR c.decotr01 = '',0,FORMAT(c.decotr01,2)) as 'DecOtros', " &_
			"FORMAT(((IF(i.tipcam01 IS NULL,0,tipcam01) * IF(i.valmer01 IS NULL,0,i.valmer01) * IF(i.factmo01 IS NULL,0,i.factmo01)) + IF(i.fletes01 IS NULL,0,i.fletes01)) +" &_
			"IF(i.segros01 IS NULL,0,i.segros01) + IF(i.incble01 IS NULL,0,i.incble01) - IF(c.decfle01 IS NULL,0,c.decfle01) - IF(c.decsgr01,0,c.decsgr01) -" &_
			"IF(c.decotr01,0,c.decotr01),0) as 'valoradu', " &_
			"IF(con1.import36 IS NOT NULL AND con1.import36 <> '',FORMAT(con1.import36,2),0) as 'Pre', " &_
			"IF(con3.import36 IS NOT NULL AND con3.import36 <> '',FORMAT(con3.import36,2),0) as 'IGI', " &_
			"IF(con4.import36 IS NOT NULL AND con4.import36 <> '',FORMAT(con4.import36,2),0) as 'DTA', " &_
			"IF(con2.import36 IS NOT NULL AND con2.import36 <> '',FORMAT(con2.import36,2),0) as 'ECI', " &_
			"IF(con7.import36 IS NOT NULL AND con7.import36 <> '',FORMAT(con7.import36,2),0) as 'OtrosImp', " &_
			"IF(con6.import36 IS NOT NULL AND con6.import36 <> '',FORMAT(con6.import36,2),0) as 'CuotCompen', " &_
			"IF(con5.import36 IS NOT NULL AND con5.import36 <> '',FORMAT(con5.import36,2),0) as 'IVA', " &_
			"(IF(con1.import36 IS NULL,0,con1.import36) + IF(con2.import36 IS NULL,0,con2.import36) + IF(con3.import36 IS NULL,0,con3.import36) + " &_
			" IF(con4.import36 IS NULL,0,con4.import36) + IF(con5.import36 IS NULL,0,con5.import36) + IF(con6.import36 IS NULL,0,con6.import36) +" &_
			" IF(con7.import36 IS NULL,0,con7.import36)) as 'TotalAduana', " &_
			"i.patent01 as 'Patente', " &_
			"conf.agente32 as 'Agente_Aduanal', " &_
			"fac.numfac39 as 'Factura', " &_
			"DATE_FORMAT(fac.fecfac39,'%d/%m/%Y') as 'FechaFac', " &_
			"TRIM(REPLACE(REPLACE(CONCAT(TRIM(prov.cvepro22),' ',TRIM(prov.nompro22)),'CANCELADO',''),'*','')) as 'Proveedor', " &_
			"i.cvepvc01 as 'Pais_Vendedor', " &_
			"fr.paiori02 as 'Pais_Origen' " &_
			"FROM rku_extranet.ssdagi01 as i " &_
			"LEFT JOIN rku_extranet.c01refer as c ON c.refe01 = i.refcia01 " &_
			"LEFT JOIN rku_extranet.ssrecp06 as rec ON rec.refcia06 = i.refcia01 " &_
			"LEFT JOIN rku_extranet.ssfact39 as fac ON i.refcia01 = fac.refcia39 " &_
			"LEFT JOIN rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02 " &_
			"LEFT JOIN rku_extranet.sscont36 as con1 ON i.refcia01 = con1.refcia36 and con1.cveimp36 = 15 and con1.fpagoi36 = 0 " &_
			"LEFT JOIN rku_extranet.sscont36 as con2 ON i.refcia01 = con2.refcia36 and con2.cveimp36 = 18 and con2.fpagoi36 = 0 " &_
			"LEFT JOIN rku_extranet.sscont36 as con3 ON i.refcia01 = con3.refcia36 and con3.cveimp36 = 6 and con3.fpagoi36 = 0 " &_
			"LEFT JOIN rku_extranet.sscont36 as con4 ON i.refcia01 = con4.refcia36 and con4.cveimp36 = 1 and con4.fpagoi36 = 0 " &_
			"LEFT JOIN rku_extranet.sscont36 as con5 ON i.refcia01 = con5.refcia36 and con5.cveimp36 = 3 and con5.fpagoi36 = 0 " &_
			"LEFT JOIN rku_extranet.sscont36 as con6 ON i.refcia01 = con6.refcia36 and con6.cveimp36 = 2 and con6.fpagoi36 = 0 " &_
			"LEFT JOIN rku_extranet.sscont36 as con7 ON i.refcia01 = con7.refcia36 and con7.fpagoi36 <> 0 " &_
			"LEFT JOIN rku_extranet.ssconf32 as conf ON i.patent01=conf.patent32 and conf.agente32 <>'' and conf.agente32 IS NOT NULL " &_
			"and i.cveadu01=conf.cveadu32 and i.cvesec01=conf.cvesec32 " &_
			"LEFT JOIN rku_extranet.ssprov22 as prov ON i.cvepro01=prov.cvepro22 " &_
			"WHERE i.cvecli01 like '" & cvecli & "' and i.firmae01 <> '' and i.firmae01 IS NOT NULL and " &_
			"i.fecpag01 >= '" & tmpAnioini & "-" & tmpmesini & "-" & tmpdiaini & "' and " &_ 
			"i.fecpag01 <= '" & tmpaniofin & "-" & tmpmesfin & "-" & tmpdiafin & "' " &_
			"group by i.refcia01, fac.numfac39, c.dmer01 " &_
			"order by CONCAT_WS('-',i.patent01,i.numped01)"

		
		Set act2= Server.CreateObject("ADODB.Recordset")
		conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=" & oficina & "_extranet; UID=pedrobm; PWD=123; OPTION=16427"

		'response.write(sqlAct)
		'Response.End
		act2.ActiveConnection = conn12
		act2.Source = sqlAct
		act2.cursortype=0
		act2.cursorlocation=2
		act2.locktype=1
		act2.open()
		
		while not act2.eof
			codigo=codigo & "<Row>"
			IF refe <> act2.fields("Referencia").value Then
				refe = act2.fields("Referencia").value
				banderita = 1
				If color = "1" then 
					color = "2"
				Else
					color = "1"
				End If
			Else
				banderita=0
			End If
			genera_html "d" & color,act2.fields("Referencia").value,"center","'String'"
			genera_html "d" & color,act2.fields("Patped").value,"center","'String'"
			genera_html "d" & color,act2.fields("Pedi").value,"center","'String'"
			genera_html "d" & color,act2.fields("Fecpag").value,"right","'String'"
			genera_html "d" & color,act2.fields("TipoOp").value,"right","'String'"
			genera_html "d" & color,act2.fields("cveped").value,"right","'String'"
			genera_html "d" & color,act2.fields("pedorg").value,"right","'String'"
			genera_html "d" & color,act2.fields("FechaR1").value,"right","'String'"
			genera_html "d" & color,act2.fields("Aduana").value,"right","'String'"
			genera_html "d" & color,act2.fields("TipoCam").value,"right","'Number'"
			genera_html "d" & color,act2.fields("Incoterm").value,"right","'String'"
			genera_html "d" & color,act2.fields("Valormext").value,"right","'Number'"
			genera_html "d" & color,act2.fields("descrip").value,"right","'String'"
			genera_html "d" & color,act2.fields("Valormn").value,"center","'Number'"
			If banderita=1 Then 
				acumvmext = acumvmext + act2.fields("Valormext").value
				acumvmn = acumvmn + act2.fields("Valormn").value
				acumincfle = acumincle + act2.fields("IncFletes").value
				acumincseg = acumincseg + act2.fields("IncSeguros").value
				acumincotr = acumincotr + act2.fields("IncOtros").value
				acumdecfle = acumdecfle + act2.fields("DecFletes").value
				acumdecseg = acumdecseg + act2.fields("DecSeguros").value
				acumdecotr = acumdecotr + act2.fields("DecOtros").value
				acumvadu = acumvadu + act2.fields("valoradu").value
				acumpre = acumpre + act2.fields("Pre").value
				acumigi = acumigi + act2.fields("IGI").value
				acumdta = acumdta + act2.fields("DTA").value
				acumeci = acumeci + act2.fields("ECI").value
				acumimp = acumimp + act2.fields("OtrosImp").value
				acumcuo = acumcuo + act2.fields("CuotCompen").value
				acumiva = acumiva + act2.fields("IVA").value
				acumtadu = acumtadu + act2.fields("TotalAduana").value

				genera_html "d" & color,act2.fields("incfletes").value,"center","'Number'"
				genera_html "d" & color,act2.fields("incseguros").value,"center","'Number'"
				genera_html "d" & color,act2.fields("incotros").value,"right","'Number'"
				genera_html "d" & color,act2.fields("decfletes"),"center","'Number'"
				genera_html "d" & color,act2.fields("DecSeguros"),"center","'Number'"
				genera_html "d" & color,act2.fields("Decotros"),"center","'Number'"
				genera_html "d" & color,act2.fields("valoradu").value,"right","'Number'"
				genera_html "d" & color,act2.fields("pre").value,"right","'Number'"
				genera_html "d" & color,act2.fields("IGI").value,"right","'Number'"
				genera_html "d" & color,act2.fields("DTA").value,"right","'Number'"
				genera_html "d" & color,act2.fields("ECI").value,"right","'Number'"
				genera_html "d" & color,act2.fields("OtrosImp").value,"right","'Number'"
				genera_html "d" & color,act2.fields("CuotCompen").value,"right","'Number'"
				genera_html "d" & color,act2.fields("IVA").value,"right","'Number'"
				genera_html "d" & color,act2.fields("TotalAduana").value,"center","'Number'"
			Else
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
				genera_html "d" & color,"0","right","'Number'"
			End If
			genera_html "d" & color,act2.fields("Patente").value,"right","'String'"
			genera_html "d" & color,act2.fields("Agente_Aduanal").value,"right","'String'"
			genera_html "d" & color,act2.fields("Factura").value,"right","'String'"
			genera_html "d" & color,act2.fields("FechaFac").value,"right","'String'"
			genera_html "d" & color,act2.fields("Proveedor").value,"right","'String'"
			genera_html "d" & color,act2.fields("Pais_Vendedor").value,"right","'String'"
			genera_html "d" & color,act2.fields("Pais_Origen").value,"right","'String'"
			'response.Write("</tr>")
			codigo=codigo &"</Row>"
			act2.movenext()
		wend
		
		codigo = codigo & "<Row>"
		genera_html "e","Total Importaciones","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "d3",acumvmext,"right","'Number'"
		'genera_ope "sum","11"
		genera_html "e","","right","'String'"
		genera_html "d3",acumvmn,"right","'Number'"
		genera_html "d3",acumincfle,"right","'Number'"
		genera_html "d3",acumincseg,"right","'Number'"
		genera_html "d3",acumincotr,"right","'Number'"
		genera_html "d3",acumdecfle,"right","'Number'"
		genera_html "d3",acumdecseg,"right","'Number'"
		genera_html "d3",acumdecotr,"right","'Number'"
		genera_html "d3",acumvadu,"right","'Number'"
		genera_html "d3",acumpre,"right","'Number'"
		genera_html "d3",acumigi,"right","'Number'"
		genera_html "d3",acumdta,"right","'Number'"
		genera_html "d3",acumeci,"right","'Number'"
		genera_html "d3",acumimp,"right","'Number'"
		genera_html "d3",acumcuo,"right","'Number'"
		genera_html "d3",acumiva,"right","'Number'"
		genera_html "d3",acumtadu,"right","'Number'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		genera_html "e","","right","'String'"
		
		codigo = codigo & "</Row>"
		
		'response.Write(codigo)
		'response.End()
		codigo = Replace((codigo), "", "a")
		codigo = Replace((codigo), "", "e")
		codigo = Replace((codigo), "", "i")
		codigo = Replace((codigo), "", "u")
		codigo = Replace((codigo), "", "u")
		'codigo = Replace((codigo), "", "A")
		'codigo = Replace((codigo), "", "E")
		'codigo = Replace((codigo), "", "I")
		'codigo = Replace((codigo), "", "O")
		'codigo = Replace((codigo), "", "U")
		'codigo = Replace((codigo), "", "")


		'codigo = Replace(codigo, ">", ")")
		'codigo = Replace(codigo, "<", "(")
		
		genera_registros = codigo
end function

sub genera_html(tipo,valor,alineacion,datatype)
	Select Case tipo
		Case "e"
			codigo = codigo & "<Cell ss:StyleID='s21'><Data ss:Type=" & datatype & ">" & valor & "</Data></Cell>"
		Case "d1"
			If datatype = "'String'" then
				codigo = codigo & "<Cell ss:StyleID='ds1'><Data ss:Type=" & datatype & ">" & valor & "</Data></Cell>"
			Else
				codigo = codigo & "<Cell ss:StyleID='dm1'><Data ss:Type=" & datatype & ">" & valor & "</Data></Cell>"
			End If
		Case "d2"
			if datatype = "'String'" then
				codigo = codigo & "<Cell ss:StyleID='ds2'><Data ss:Type=" & datatype & ">" & valor & "</Data></Cell>"
			Else
				codigo = codigo & "<Cell ss:StyleID='dm2'><Data ss:Type=" & datatype & ">" & valor & "</Data></Cell>"
			End If
		case "d3"
			codigo = codigo & "<Cell ss:StyleID='ft1'><Data ss:Type=" & datatype & ">" & valor & "</Data></Cell>"
	End Select
end sub

sub genera_ope(tipo,columna)
	Select Case tipo
		Case "sum"
			'codigo = codigo & "<Cell ss:StyleID='s21'><Data ss:Type='String'>PRUEBA EXITOSA</Data></Cell>"
			'Response.Write("<Cell ss:StyleID='s21 ss:Formula=" & Chr(34) & "=SUM(R2C11:R[-1]C)" & Chr(34) & "></Cell>)")
			'Response.End
			'codigo = codigo & "<Cell ss:StyleID='s21 ss:Formula=" & Chr(34) & "=SUM(R2C11:R[-1]C)" & Chr(34) & "></Cell>)"
	End Select
End Sub
%>
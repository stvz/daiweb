<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000
dim codigo, genera_registros
genera_registros= ""

strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

if not permi = "" then
	permi = "  and (" & permi & ") "
end if
AplicaFiltro = False
strFiltroCliente = ""
'strFiltroCliente = request.Form("txtCliente")


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

	nocolumns = 0
	tablamov = ""
	nocolumns=30
	
	For xtipo = 0 To 0
        If xtipo = 0 then
		 tablamov = "ssdagi01"
			For y = 0 To 4
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
							query = query & " union all "
							else
								if y = 4 then
								strOficina = "lzr"
								query = query & GeneraSQL & " order by 2,ordfra02 "
								End if
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
		Response.Write("No hay datos para esas condiciones" & query)
	Else
		if Tiporepo = 2 Then

		End If
		
		
		
			codigo = ""
			 
			 codigo=codigo &"<Row>"
				 genera_html "e","PEDIMENTO","center"								
				 genera_html "e","REFERENCIA","center"
				 genera_html "e","FRACCION","center"
				 genera_html "e","DESCRIPCION","center"
				 genera_html "e","FACTURAS","center"
				 genera_html "e","HOUSE BL","center"
				 genera_html "e","P/O No","center"
				 genera_html "e","INCOTERMS","center"
				 genera_html "e","FECHA DE ENTRADA","center"
				 genera_html "e","TIPO DE CAMBIO","center"
				 genera_html "e","FACTOR MONEDA","center"
				 genera_html "e","VALOR FACTURA ME","center"
				 genera_html "e","VALOR MERCANCIA MON NAC","center"
				 genera_html "e","VALOR DOLARES","center"
				 genera_html "e","TOTAL QUANTITY","center"
				 genera_html "e","INVOICE AMOUNT","center"
				 genera_html "e","INVOICE CURRENCY","center"
				 genera_html "e","MATERIAL NUMBER","center"
				 genera_html "e","VALOR ADUANA","center"
				 genera_html "e","DTA","center"
				 genera_html "e","DTA (CALC)","center"
				 genera_html "e","IGI","center"
				 genera_html "e","IGI (CALC)","center"
				 genera_html "e","PRV","center"
				 genera_html "e","IVA","center"
				 genera_html "e","IVA (CALC)","center"
				 genera_html "e","TOTAL IMPUESTOS","center"
				 genera_html "e","TOTAL IMPUESTOS (CALC)","center"
				 genera_html "e","CVE TIPO TASA DTA","center"
				 genera_html "e","TASA ADV","center"
			 codigo=codigo &"</Row>"

	
		
		datos = ""
		contador=4
		nueva=true
		refcia = ""
		Do Until RSops.EOF
			contador = contador + 1
			
			if refcia <> "" then
				if RSops.Fields.Item("Referencia").Value = refcia  then
					nueva = false
				else
					nueva = true
				end if
			end if
			
			if nueva then
				prv = RSops.Fields.Item("PRV").Value
				dta = RSops.Fields.Item("DTA").Value
				totimp = RSops.Fields.Item("TotalImpuestos").Value
				dolares = RSops.Fields.Item("Valor Dolares").Value
				
			else
				prv = 0
				dta = 0
				totimp = 0
				dolares = 0
				
			end if
			
			
			codigo = codigo &"<Row>"
			
			genera_html "d",RSops.Fields.Item("Pedimento").Value, "center"
			genera_html "d",RSops.Fields.Item("Referencia").Value, "center"
			genera_html "d",RSops.Fields.Item("FraccionAranc").Value, "center"
			genera_html "d",RSops.Fields.Item("Descripcion").Value, "center"
			genera_html "d",RSops.Fields.Item("Facturas").Value, "center"
			genera_html "d",RSops.Fields.Item("House B/L").Value, "center"
			genera_html "d",RSops.Fields.Item("P/O No").Value, "center"
			genera_html "d",RSops.Fields.Item("Fecha de Entrada").Value, "center"
			genera_html "d",RSops.Fields.Item("Tipo de Cambio").Value, "center"
			genera_html "d",RSops.Fields.Item("Factor Moneda").Value, "right"
			genera_html "d",oo(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value), "center"
			genera_html "d",RSops.Fields.Item("Valor Merc Mon Nac").Value, "center"
			genera_html "d",dolares, "center"
			genera_html "d",RSops.Fields.Item("Total Quantity").Value, "center"
			genera_html "d",RSops.Fields.Item("Invoice Amount").Value, "center"
			genera_html "d",RSops.Fields.Item("Invoice Currency").Value, "center"
			genera_html "d",RSops.Fields.Item("Material Number").Value, "center"
			genera_html "d",RSops.Fields.Item("Valor Aduana").Value, "center"
			genera_html "d",dta, "center"
			genera_html "d","=TEXTO(REDONDEAR(SI(AC" & cstr(contador) & "=7,0.008*S" & cstr(contador) & ",SI(AC" & cstr(contador) & "=4,T" & cstr(contador) & ",0)),0), " & """" & "#,##0" & """" & ")", "center"
			genera_html "d",RSops.Fields.Item("IGI").Value, "center"
			genera_html "d","=TEXTO(REDONDEAR((S"&cstr(contador) & "* ( AD"&cstr(contador) & ")/100),0), " & """" & "#,##0" & """" & ")", "center"
			genera_html "d",prv, "center"
			genera_html "d",RSops.Fields.Item("IVA").Value, "center"
			genera_html "d","=TEXTO(REDONDEAR(((S"&cstr(contador) & "+U"&cstr(contador) & "+W"&cstr(contador) & ")*0.16),0), " & """" & "#,##0" & """" & ")", "center"
			genera_html "d",totimp, "center"
			genera_html "d","=TEXTO(REDONDEAR(U"&cstr(contador) & "+W"&cstr(contador) & "+X"&cstr(contador)& "+Z"&cstr(contador) & ",0), " & """" & "#,##0" & """" & ")", "center"
			genera_html "d",RSops.Fields.Item("tt_dta01").Value, "center"
			genera_html "d",RSops.Fields.Item("tasadv02").Value, "center"

			refcia =  RSops.Fields.Item("Referencia").Value
			
			' response.Write("</tr>")
			codigo=codigo &"</Row>"
			Rsops.MoveNext()
		Loop
		
			codigo = Replace((codigo), "á", "a")
			codigo = Replace((codigo), "é", "e")
			codigo = Replace((codigo), "í", "i")
			codigo = Replace((codigo), "ó", "u")
			codigo = Replace((codigo), "ú", "u")
 	 genera_registros = codigo
	 
	 


	End If
end if

function GeneraSQL
	SQL = ""
	condicion = filtro
	SQL = 	"SELECT cast(CONCAT_WS(' ', MID(i.fecpag01,3,2), i.cveadu01, i.patent01, i.numped01) as char) Pedimento," &_
			" i.refcia01 Referencia, " &_
			" fr.fraarn02 FraccionAranc, " &_
			" fr.d_mer102 Descripcion, " &_
			" group_concat(distinct f.numfac39,char(05) separator '') 'Facturas', " &_
			" (select group_concat(gui1.numgui04 ) from " & strOficina & "_extranet.ssguia04 as gui1 where gui1.refcia04 = i.refcia01 and gui1.patent04 = i.patent01 and gui1.adusec04 = i.adusec01) 'House B/L', " &_
			" ar.pedi05  as'P/O No', " &_
			" group_concat(distinct f.terfac39) Incoterms, " &_
			" i.fecent01 as 'Fecha de Entrada', " &_
			" i.tipcam01 as 'Tipo de Cambio', " &_
			" cast(group_concat(distinct f.facmon39) as char) 'Factor Moneda', " &_
		
			" round(sum(fr.prepag02),2) 'Valor Merc Mon Nac', " &_
			" format(i.valdol01,2) as 'Valor Dolares', " &_
			" format(sum(ar.caco05),0) as 'Total Quantity', " &_
			" format(sum(ar.vafa05),2) as 'Invoice Amount', " &_
			" group_concat(distinct f.monfac39) as 'Invoice Currency', " &_
			" group_concat(ar.item05) as 'Material Number', " &_
			" (select format(sum(fra.vaduan02),0) from " & strOficina & "_extranet.ssfrac02 as fra where i.refcia01 = fra.refcia02 and fra.fraarn02 = fr.fraarn02 ) 'Valor Aduana', " &_
			" format(cf1.import36,0) DTA, " &_
			" format(fr.i_adv102 + fr.i_adv202 + fr.i_adv302,0) IGI, " &_
			" format(cf15.import36,0) PRV, " &_
			" (select format(sum(fra.i_iva102 + fra.i_iva202 + fra.i_iva302),0) from " & strOficina & "_extranet.ssfrac02 as fra where i.refcia01 = fra.refcia02 and fra.fraarn02 = fr.fraarn02 ) 'IVA', " &_
			" format((ifnull(cf1.import36,0)  + ifnull(cf3.import36,0) + ifnull(cf6.import36,0) + ifnull(cf15.import36,0) ),0) TotalImpuestos, " &_
			" i.tt_dta01, " &_
			" fr.ordfra02, " &_
			" fr.tasadv02, " &_
			" i.adusec01 as adu, " &_
			" i.patent01 " &_
			
			"from " & strOficina & "_extranet." & tablamov & " as i " &_
			" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
			"  left join " & strOficina & "_extranet.c01refer as r on r.refe01 = i.refcia01 " &_
			"   LEFT join " & strOficina & "_extranet.d01conte as ct on ct.refe01 = i.refcia01 " &_
			"    inner join " & strOficina & "_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01 " &_
			"     left join " & strOficina & "_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " &_
			"      left join " & strOficina & "_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05 " &_
			"           left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1' and cf1.adusec36 = i.adusec01 and cf1.patent36 =i.patent01  " &_
			"            left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3' and cf3.adusec36 = i.adusec01 and cf3.patent36 =i.patent01" &_
			"             left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6' and cf6.adusec36 = i.adusec01 and cf6.patent36 =i.patent01 " &_
			"              left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15' and cf15.adusec36 = i.adusec01 and cf15.patent36 =i.patent01 " &_
			"where cc.rfccli18 = 'SEM950215S98' " & condicion &_
			" group by i.refcia01, fr.fraarn02 "
			
	GeneraSQL = SQL
end function


function filtro
	cadena = "'" & replace (Request.Form("CServisWeb"),",","','") & "'"
	'condicion = " and cc.rfccli18 = 'SEM950215S98' and i.refcia01 in( " & Request.Form("CServisWeb") & ")"
	condicion = " and i.refcia01 in (" & cadena & ")"
	filtro = condicion
end function

function oo(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
 
sqlAct=" select group_concat(format(fact.valmex39,2) separator '|') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "' and fact.numfac39 in (select  arti.fact05 from " & oficina & "_extranet.d05artic as arti where arti.refe05 = '" & referencia & "' and arti.frac05 = '"&fraccion&"')"
'		response.write (sqlact)
'		response.end()
Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		oo =act2.fields("val").value
	else
		oo =valor
	end if
	

end function


sub genera_html(tipo,valor,alineacion)

 if(tipo = "e")then
  'response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
  codigo = codigo&"<Cell ss:StyleID='s21'><Data ss:Type='String'>"&valor&"</Data></Cell>" 
 else 
 '  response.Write("<td align="&c&alineacion&c&" nowrap><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
  codigo = codigo&"<Cell><Data ss:Type='String'>"&valor&"</Data></Cell>"
 end if

end sub


%>

<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE IMPUESTOS SAMSUNG.... ::</TITLE>
	</HEAD>
	<BODY>
		<%=genera_registros
			Dim XML : XML = "<?xml version='1.0'?>" & _
              "<?mso-application progid='Excel.Sheet'?>" & _
               "<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' " & _
               " xmlns:o='urn:schemas-microsoft-com:office:office' "& _
               " xmlns:x='urn:schemas-microsoft-com:office:excel' " & _
               " xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet' " & _
               " xmlns:html='http://www.w3.org/TR/REC-html40'>" & _
               " <DocumentProperties xmlns='urn:schemas-microsoft-com:office:office'>" & _
               "    <Author>Pedro Bautista</Author>" & _
               "    <LastAuthor>Pedro Bautista</LastAuthor>" & _
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
			   " <Style ss:ID='s21'> " & _
			   " <Interior ss:Color='#330099' ss:Pattern='Solid'></Interior> " & _
			   "  <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='12' ss:Color='#FFFFFF'/>" & _
			   "  </Style> " & _
               "</Styles>" & _
               "<Worksheet ss:Name='ENC IMPO'>" & _
                        " <Table x:FullColumns='1' " & _
						" x:FullRows='1' ss:DefaultRowHeight='15'>" & _
						 genera_registros & _
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
           "           <ActiveRow>16</ActiveRow>" & _
           "           <ActiveCol>8</ActiveCol>" & _
           "       </Pane>" & _
           "    </Panes>" & _
           "    <ProtectObjects>False</ProtectObjects>" & _
           "    <ProtectScenarios>False</ProtectScenarios>" & _
           "</WorksheetOptions>" & _
           "</Worksheet>" & _
           "<Worksheet ss:Name='FACTURAS IMPO'>" & _
						"    <Table x:FullColumns='1' " & _
						"      x:FullRows='1' ss:DefaultRowHeight='15'>" & _
					   genera_registros & _
						 "</Table>" & _
           " <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "   <PageSetup>" & _
           "       <Header x:Margin='0.3'/>" & _
           "       <Footer x:Margin='0.3'/>" & _
           "       <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "    </PageSetup>" & _
           "   <ProtectObjects>False</ProtectObjects>" & _
           "   <ProtectScenarios>False</ProtectScenarios>" & _
           "  </WorksheetOptions>" & _
           " </Worksheet>" & _
           " <Worksheet ss:Name='ENC EXPO'>" & _
						"    <Table x:FullColumns='1' " & _
						"      x:FullRows='1' ss:DefaultRowHeight='15'>" & _
						 genera_registros & _
						"</Table>" & _
           "  <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "   <PageSetup>" & _
           "    <Header x:Margin='0.3'/>" & _
           "    <Footer x:Margin='0.3'/>" & _
           "    <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "   </PageSetup>" & _
           "   <ProtectObjects>False</ProtectObjects>" & _
           "   <ProtectScenarios>False</ProtectScenarios>" & _
           "  </WorksheetOptions>" & _
           " </Worksheet>" & _
		   " <Worksheet ss:Name='FACTURAS EXPO'>" & _
						"    <Table x:FullColumns='1' " & _
						"      x:FullRows='1' ss:DefaultRowHeight='15'>" & _
						 genera_registros & _
						"</Table>" & _
           "  <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "   <PageSetup>" & _
           "    <Header x:Margin='0.3'/>" & _
           "    <Footer x:Margin='0.3'/>" & _
           "    <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "   </PageSetup>" & _
           "   <ProtectObjects>False</ProtectObjects>" & _
           "   <ProtectScenarios>False</ProtectScenarios>" & _
           "  </WorksheetOptions>" & _
           " </Worksheet>" & _
           "</Workbook>"
		   
		   		With Response
				   .Charset = "UTF-8"
				   .Clear
				   .ContentType = "excel/ms-excel"
				   .AddHeader "Content-Disposition","attachment; filename=Rep_Imp_Sam.xls"
				   .AddHeader "Content-Length", Len(XML)
				   .Write XML
				   .Flush
				   .End
				End With
		%>
	</BODY>
</HTML>

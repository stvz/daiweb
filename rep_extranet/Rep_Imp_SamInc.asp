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
	IF  False Then
		Response.Write(query)
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
												"GRUPO ZEGO" &_
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
												" SOLICITUD DE IMPUETOS" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td>" &_
								"</td>" &_
							"</tr>" &_
				"</table>"
		
		header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">" &_
							    celdahead("PEDIMENTO") &_
								celdahead("REFERENCIA") &_
								celdahead("FRACCION") &_
								celdahead("DESCRIPCION") &_
								celdahead("FACTURAS") &_
								celdahead("HOUSE BL") &_
								celdahead("P/O No") &_
								celdahead("INCOTERMS") &_
								celdahead("FECHA DE ENTRADA") &_
								celdahead("TIPO DE CAMBIO") &_
								celdahead("FACTOR MONEDA") &_
								celdahead("VALOR FACTURA ME") &_
								celdahead("VALOR MERCANCIA MON NAC") &_
								celdahead("VALOR DOLARES") &_
								celdahead("TOTAL QUANTITY") &_
								celdahead("INVOICE AMOUNT") &_
								celdahead("INVOICE CURRENCY") &_
								celdahead("MATERIAL NUMBER") &_
								celdahead("VALOR ADUANA") &_
								celdahead("DTA") &_
								celdahead("DTA (CALC)") &_
								celdahead("IGI") &_
								celdahead("IGI (CALC)") &_
								celdahead("PRV") &_
								celdahead("IVA") &_
								celdahead("IVA (CALC)") &_
								celdahead("TOTAL IMPUESTOS") &_
								celdahead("TOTAL IMPUESTOS (CALC)") &_
								celdahead("CVE TIPO TASA DTA") &_
								celdahead("TASA ADV")
		header = header &	"</tr>"
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
			
			
			datos = datos &	"<tr>" &_
			celdadatos(RSops.Fields.Item("Pedimento").Value) &_
			celdadatos(RSops.Fields.Item("Referencia").Value) &_
			celdadatos(RSops.Fields.Item("FraccionAranc").Value) &_
			celdadatos(RSops.Fields.Item("Descripcion").Value) &_
			celdadatos(FACTURAS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(RSops.Fields.Item("House B/L").Value) &_
			celdadatos(PO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(INCOTERMS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(RSops.Fields.Item("Fecha de Entrada").Value) &_
			celdadatos(RSops.Fields.Item("Tipo de Cambio").Value) &_
			celdadatos(FACMON(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(VFME(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdanumero(RSops.Fields.Item("Valor Merc Mon Nac").Value) &_
			celdanumero(dolares) &_
			celdanumeroentero(RSops.Fields.Item("Total Quantity").Value) &_
			celdanumero(RSops.Fields.Item("Invoice Amount").Value) &_
			celdadatos(MONFAC(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(MAN(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdanumeroentero(RSops.Fields.Item("Valor Aduana").Value) &_
			celdanumeroentero(dta) &_
			celdanumero("=SI(AC" & cstr(contador) & "=7,0.008*S" & cstr(contador) & ",SI(AC" & cstr(contador) & "=4,T" & cstr(contador) & ",0))") &_
			celdanumeroentero(RSops.Fields.Item("IGI").Value) &_
			celdanumero("=REDONDEAR((S"&cstr(contador) & "* ( AD"&cstr(contador) & ")/100),0)") &_
			celdanumeroentero(prv) &_
			celdanumeroentero(RSops.Fields.Item("IVA").Value) &_
			celdanumero("=REDONDEAR(((S"&cstr(contador) & "+U"&cstr(contador) & "+W"&cstr(contador) & ")*0.16),0)") &_
			celdanumeroentero(totimp) &_
			celdanumero("=REDONDEAR(U"&cstr(contador) & "+W"&cstr(contador) & "+X"&cstr(contador)& "+Z"&cstr(contador) & ",0)") &_
			celdadatos(RSops.Fields.Item("tt_dta01").Value) &_
			celdadatos(RSops.Fields.Item("tasadv02").Value)
			datos = datos &	"</tr>"
		
			refcia =  RSops.Fields.Item("Referencia").Value
			
			Rsops.MoveNext()
		Loop
		
		sumas = ""
		sumas = "<tr>" &_
		"<td colspan=""" & 11 & """>" &_
			"<center>" &_
						"" &_
			"</center>" &_
		"</td>" &_
								
		celdasumas("SUMAS") &_
		celdasumasnumero("=SUMA(M5:M"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(N5:N"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(O5:O"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(P5:P"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(S5:S"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(T5:T"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(U5:U"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(V5:V"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(W5:W"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(X5:X"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Y5:Y"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Z5:Z"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AA5:AA"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AB5:AB"&cstr(contador)&")") &_
		celdadatos("")
		sumas =  sumas & "</tr>"
	 
	   
 	 html = info & header & datos & sumas & "</table><br>" 

	End If
end if

function GeneraSQL
	SQL = ""
	condicion = filtro
	SQL = 	"SELECT cast(CONCAT_WS(' ', MID(i.fecpag01,3,2), i.cveadu01, i.patent01, i.numped01) as char) Pedimento," &_
			" i.refcia01 Referencia, " &_
			" fr.fraarn02 FraccionAranc, " &_
			" group_concat(fr.d_mer102) Descripcion, " &_
			" (select group_concat(gui1.numgui04 ) from " & strOficina & "_extranet.ssguia04 as gui1 where gui1.refcia04 = i.refcia01 and gui1.patent04 = i.patent01 and gui1.adusec04 = i.adusec01) 'House B/L', " &_
			" i.fecent01 as 'Fecha de Entrada', " &_
			" i.tipcam01 as 'Tipo de Cambio', " &_
			" round(sum(fr.prepag02),2) 'Valor Merc Mon Nac', " &_
			" format(i.valdol01,2) as 'Valor Dolares', " &_
			" sum(fr.cancom02) as 'Total Quantity', " &_
			" sum(fr.vmerme02) as 'Invoice Amount', " &_
			" format(sum(fr.vaduan02),0) 'Valor Aduana', " &_
			" format(cf1.import36,0) DTA, " &_
			" format(fr.i_adv102 + fr.i_adv202 + fr.i_adv302,0) IGI, " &_
			" format(cf15.import36,0) PRV, " &_
			" format(sum(fr.i_iva102 + fr.i_iva202 + fr.i_iva302),0) 'IVA', " &_
			" format((ifnull(cf1.import36,0)  + ifnull(cf3.import36,0) + ifnull(cf6.import36,0) + ifnull(cf15.import36,0) ),0) TotalImpuestos, " &_
			" i.tt_dta01, " &_
			" fr.ordfra02, " &_
			" fr.tasadv02, " &_
			" i.adusec01 as adu, " &_
			" i.patent01 " &_
			
			"from " & strOficina & "_extranet." & tablamov & " as i " &_
			" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
			"      left join " & strOficina & "_extranet.ssfrac02 as fr on fr.refcia02 = i.refcia01  and fr.patent02 = i.patent01 and fr.adusec02 = i.adusec01  " &_
			"           left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1' and cf1.adusec36 = i.adusec01 and cf1.patent36 =i.patent01  " &_
			"            left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3' and cf3.adusec36 = i.adusec01 and cf3.patent36 =i.patent01" &_
			"             left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6' and cf6.adusec36 = i.adusec01 and cf6.patent36 =i.patent01 " &_
			"              left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15' and cf15.adusec36 = i.adusec01 and cf15.patent36 =i.patent01 " &_
			"where cc.rfccli18 = 'SEM950215S98' " & condicion &_
			" group by i.refcia01, fr.fraarn02 "
				
	GeneraSQL = SQL
end function


function celdahead(texto)
	
	'textodos = texto
	
	'if isnumeric(texto) then
	' texto = formatNumber(textodos,2)
	'end if
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
	'If IsNull(texto) = True Or texto = "" Then
	'	texto = "-"
	'End If
	cell = 	"<td align=""center"">" &_
				
					texto &_
			
			"</td>"
	celdadatos = cell
end function

function celdanumero(texto)
	'If IsNull(texto) = True Or texto = "" Then
	'	texto = "-"
	'End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumero = cell
end function

function celdanumeroentero(texto)
	'If IsNull(texto) = True Or texto = "" Then
	'	texto = "-"
	'End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumeroentero = cell
end function


function celdasumas(texto)
	'If IsNull(texto) = True Or texto = "" Then
	'	texto = "-"
	'End If
	'"<font color=""#FFFFFG"">" &_
	'"</font>" &_
	cell = 	"<td align=""center"" style=""font-weight: bold"" >" &_
				
					texto &_
	
			"</td>"
	celdasumas = cell
end function

function celdasumasnumero(texto)
	'If IsNull(texto) = True Or texto = "" Then
	'	texto = "-"
	'End If
	'"<font color=""#FFFFFG"">" &_
	'"</font>" &_
	cell = 	"<td align=""center"" style=""font-weight: bold"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
	
			"</td>"
	celdasumasnumero = cell
end function

function celdasumasnumeroentero(texto)
	'If IsNull(texto) = True Or texto = "" Then
	'	texto = "-"
	'End If
	'"<font color=""#FFFFFG"">" &_
	'"</font>" &_
	cell = 	"<td align=""center"" style=""font-weight: bold"" style=""mso-number-format:'#,##0';"">" &_
				
					texto &_
	
			"</td>"
	celdasumasnumeroentero = cell
end function

function filtro
	'if Vckcve = 0 then
	'	condicion = " and cc.rfccli18 = '" & Vrfc & "' "
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
	'end if
	cadena = "'" & replace (Request.Form("CServisWeb"),",","','") & "'"
	'condicion = " and cc.rfccli18 = 'SEM950215S98' and i.refcia01 in( " & Request.Form("CServisWeb") & ")"
	condicion = " and i.refcia01 in (" & cadena & ")"
	filtro = condicion
end function

function VFME(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
  if (ucase(oficina) = "PAN")then
	oficina = "DAI"
 end if

sqlAct=" select group_concat(distinct format(fact.valmex39,2) separator '|') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"

'sqlAct=" select group_concat(format(fact.valmex39,2) separator '|') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
'		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "' and fact.numfac39 in (select  arti.fact05 from " & oficina & "_extranet.d05artic as arti where arti.refe05 = '" & referencia & "' and arti.frac05 = '"&fraccion&"')"
'		response.write (sqlact)
'		response.end()

'"select group_concat(fac.valmex39) as val from "&oficina&"_extranet.ssfact39 fac where fac.numfac39 in"&_
'"(select fact.numfac39 from "&oficina&"_extranet.ssfact39 fact "&_
'" left join "&oficina&"_extranet.d05artic as art on art.refe05 = '"&referencia&"' and art.fact05 = fact.numfac39 "&_
'" left join "&oficina&"_extranet.ssfrac02 as frt on frt.refcia02 = '"&referencia&"' and frt.fraarn02 = art.frac05 and frt.ordfra02 = art.agru05 "&_
'" where fact.refcia39 =  '"&referencia&"' and fact.adusec39 = '"&aduana&"' and fact.patent39 = '"&patente&"' and frt.fraarn02 = "&fraccion&_
'")"


			'sqlAct2=" cast(group_concat(distinct f.facmon39) as char) 'Factor Moneda', " &_
			'" cast((select group_concat(distinct concat(fact.numfac39,':',fact.valmex39),char(05) separator '') from " & strOficina & "_extranet.ssfact39 fact " &_
			'" inner join " & strOficina & "_extranet.d05artic as art on art.refe05 = fact.refcia39 and art.fact05 = fact.numfac39 " &_
			'" inner join " & strOficina & "_extranet.ssfrac02 as frt on frt.refcia02 = fact.refcia39 and frt.fraarn02 = art.frac05 and frt.ordfra02 = art.agru05 " &_
			'" where  fact.refcia39 = i.refcia01 and fact.adusec39 = i.adusec01 and fact.patent39 = i.patent01 and frt.fraarn02 = fr.fraarn02)as char) 
			
Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		VFME =act2.fields("val").value
	else
		VFME =valor
	end if
end function

function PO(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
  if (ucase(oficina) = "PAN")then
	oficina = "DAI"
 end if
 
sqlAct=" select group_concat(distinct ar.pedi05,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as f " &_
	" left join " & oficina & "_extranet.d05artic as ar on ar.refe05 = f.refcia39 and  ar.fact05 =f.numfac39  and ar.frac05 = " & fraccion &_
	" where f.refcia39 = '" & referencia & "' and  f.adusec39 = '" & aduana & "' and f.patent39 = '" & patente & "'"
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
		PO =act2.fields("val").value
	else
		PO =valor
	end if
end function


function MAN(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
  if (ucase(oficina) = "PAN")then
	oficina = "DAI"
 end if
 
sqlAct=" select group_concat(ar.item05) as val from " & oficina & "_extranet.ssfact39 as f " &_
	" left join " & oficina & "_extranet.d05artic as ar on ar.refe05 = f.refcia39 and  ar.fact05 =f.numfac39  and ar.frac05 = " & fraccion &_
	" where f.refcia39 = '" & referencia & "' and  f.adusec39 = '" & aduana & "' and f.patent39 = '" & patente & "'"
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
		MAN =act2.fields("val").value
	else
		MAN =valor
	end if
end function

function FACTURAS(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
  if (ucase(oficina) = "PAN")then
	oficina = "DAI"
 end if
sqlAct=" select group_concat(distinct fact.numfac39,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"
 
'sqlAct=" select group_concat(fact.numfac39,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
'		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "' and fact.numfac39 in (select  arti.fact05 from " & oficina & "_extranet.d05artic as arti where arti.refe05 = '" & referencia & "' and arti.frac05 = '"&fraccion&"')"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		FACTURAS =act2.fields("val").value
	else
		FACTURAS =valor
	end if
end function

function INCOTERMS(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
 if (ucase(oficina) = "ALC")then
	oficina = "LZR"
 end if
  if (ucase(oficina) = "PAN")then
	oficina = "DAI"
 end if
 
 sqlAct=" select group_concat(distinct fact.terfac39,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"

'sqlAct=" select group_concat(fact.terfac39,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
'		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "' and fact.numfac39 in (select  arti.fact05 from " & oficina & "_extranet.d05artic as arti where arti.refe05 = '" & referencia & "' and arti.frac05 = '"&fraccion&"')"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		INCOTERMS =act2.fields("val").value
	else
		INCOTERMS =valor
	end if
end function

function FACMON(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
if (ucase(oficina) = "ALC")then
	oficina = "LZR"
end if
if (ucase(oficina) = "PAN")then
	oficina = "DAI"
end if

sqlAct=" select cast(group_concat(distinct fact.facmon39) as char) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"
 
'sqlAct=" select cast(group_concat(distinct fact.facmon39) as char) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
'		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "' and fact.numfac39 in (select  arti.fact05 from " & oficina & "_extranet.d05artic as arti where arti.refe05 = '" & referencia & "' and arti.frac05 = '"&fraccion&"')"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		FACMON =act2.fields("val").value
	else
		FACMON =valor
	end if
end function

function MONFAC(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="0"
 
if (ucase(oficina) = "ALC")then
	oficina = "LZR"
end if
if (ucase(oficina) = "PAN")then
	oficina = "DAI"
end if

sqlAct=" select group_concat(distinct fact.monfac39) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"

'sqlAct=" select group_concat(fact.monfac39) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
'		" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "' and fact.numfac39 in (select  arti.fact05 from " & oficina & "_extranet.d05artic as arti where arti.refe05 = '" & referencia & "' and arti.frac05 = '"&fraccion&"')"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		MONFAC =act2.fields("val").value
	else
		MONFAC =valor
	end if
end function
%>

<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE IMPUESTOS SAMSUNG.... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html
		%>
	</BODY>
</HTML>

<!--
			'Dim ApExcel 
			'Set ApExcel = CreateObject("Excel.application")
			'ApExcel.Visible = True
			'ApExcel.Workbooks.open("C:\Users\alanaci\Desktop\SAMSUNG\Rep_Imp_Sam4.xls")
			'ApExcel.Range("AB5:AB8").Select
			'ApExcel.Selection.NumberFormat = "#,##0.00"
-->
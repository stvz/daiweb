<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">

<%Server.ScriptTimeout=15000
Dim sRuta, codigo
sRuta = ""

fechaactualizada = ""


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
	
	if Request.Form("Enviar") = "t" then
		if Request.Form("txtCorreo") = "" then
			Response.Write("<strong><br><font color=""#006699"" size=""4"" face=""Arial, Helvetica, sans-serif"">Debe de escribir por lo menos un correo.<br> Gracias.</font></strong>")
			Response.End()	
		end if
	end if
	if checaCargas then
		Response.Write("<strong><br><font color=""#006699"" size=""4"" face=""Arial, Helvetica, sans-serif"">Las Bases de Datos se estan actualizando y no es posible llevar a cabo su solicitud. <br> Por Favor intente de nuevo en unos momentos. <br> Gracias.</font></strong>")
		Response.End()
	end if
	
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
		'response.write (query)
		'response.end()
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	Set RSops = CreateObject("ADODB.RecordSet")

	Set RSops = ConnStr.Execute(query)
	IF  False Then
		Response.Write(query)
	Else
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment; filename=Reporte_Impuestos.xls"
			'Response.ContentType = "excel/ms-excel"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		fechaactualizada = FechasAct()
		
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
								celdahead("MATERIAL NUMBER") &_
								celdahead("FACTURAS") &_
								celdahead("HOUSE BL o CONTENEDOR(LAZ)") &_
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
								celdahead("VALOR TOTAL FACT MON NAC") &_
								celdahead("INVOICE CURRENCY") &_
								celdahead("VALOR SEGUROS PED") &_
								celdahead("VALOR SEGUROS FRACCION") &_
								celdahead("SEGUROS PED") &_
								celdahead("SEGUROS FRACCION") &_
								celdahead("FLETES PED") &_
								celdahead("FLETES FRACCION") &_
								celdahead("EMBALAJES PED") &_
								celdahead("EMBALAJES FRACCION") &_
								celdahead("OTROS INCREMENTABLES PED") &_
								celdahead("OTROS INC FRACCION") &_
								celdahead("VALOR ADUANA") &_
								celdahead("VALOR ADUANA CALC") &_
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
			else
				prv = 0
				dta = 0
				totimp = 0
			end if
			
			
			datos = datos &	"<tr>" &_
			celdadatos(RSops.Fields.Item("Pedimento").Value) &_
			celdadatos(RSops.Fields.Item("Referencia").Value) &_
			celdadatos(RSops.Fields.Item("FraccionAranc").Value) &_
			celdadatos(RSops.Fields.Item("Descripcion").Value) &_
			celdadatos(MAN(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(FACTURAS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(conte(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(PO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(INCOTERMS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(RSops.Fields.Item("Fecha de Entrada").Value) &_
			celdadatos(RSops.Fields.Item("Tipo de Cambio").Value) &_
			celdadatos(FACMON(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdadatos(VFME(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdanumero("=REDONDEAR(O" & cstr(contador) & "*K" & cstr(contador) & ",0)") &_
			celdanumero("=Q" & cstr(contador) & "*L" & cstr(contador)) &_
			celdanumeroentero(RSops.Fields.Item("Total Quantity").Value) &_
			celdanumero(RSops.Fields.Item("Invoice Amount").Value) &_
			celdanumero(RSops.Fields.Item("Tot Fac Mon Nac").Value) &_
			celdadatos(MONFAC(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)) &_
			celdanumero(RSops.Fields.Item("Valor Seguros").Value) &_
			celdanumero("=(N" & cstr(contador) & "/R" & cstr(contador) & ")*T" & cstr(contador)) &_
			celdanumero(RSops.Fields.Item("Seguros").Value) &_
			celdanumero("=(N" & cstr(contador) & "/R" & cstr(contador) & ")*V" & cstr(contador)) &_
			celdanumero( RSops.Fields.Item("Fletes").Value) &_
			celdanumero("=(N" & cstr(contador) & "/R" & cstr(contador) & ")*X" & cstr(contador)) &_
			celdanumero(RSops.Fields.Item("Embalajes").Value) &_
			celdanumero("=(N" & cstr(contador) & "/R" & cstr(contador) & ")*Z" & cstr(contador)) &_
			celdanumero(RSops.Fields.Item("OtrosInc").Value) &_
			celdanumero("=(N" & cstr(contador) & "/R" & cstr(contador) & ")*AB" & cstr(contador)) &_
			celdanumeroentero(RSops.Fields.Item("Valor Aduana").Value) &_
			celdanumero("=REDONDEAR(N"&cstr(contador) & "+W"&cstr(contador) & "+Y"&cstr(contador)& "+AA"&cstr(contador) & "+AC"&cstr(contador)& ",0)") &_
			celdanumeroentero(dta) &_
			celdanumero("=SI(AO" & cstr(contador) & "=7,0.008*AD" & cstr(contador) & ",SI(AO" & cstr(contador) & "=4,af" & cstr(contador) & ",0))") &_
			celdanumeroentero(RSops.Fields.Item("IGI").Value) &_
			celdanumero("=REDONDEAR((AD"&cstr(contador) & "* (AP"&cstr(contador) & ")/100),0)") &_
			celdanumeroentero(prv) &_
			celdanumeroentero(RSops.Fields.Item("IVA").Value) &_
			celdanumero("=REDONDEAR(((AD"&cstr(contador) & "+AG"&cstr(contador) & "+AI"&cstr(contador) & ")*0.16),0)") &_
			celdanumeroentero(totimp) &_
			celdanumero("=REDONDEAR(AG"&cstr(contador) & "+AI"&cstr(contador) & "+AJ"&cstr(contador)& "+AL"&cstr(contador) & ",0)") &_
			celdadatos(RSops.Fields.Item("tt_dta01").Value) &_
			celdadatos(RSops.Fields.Item("tasadv02").Value)
			datos = datos &	"</tr>"
		
			refcia =  RSops.Fields.Item("Referencia").Value

			Rsops.MoveNext()
		Loop
		
		sumas = ""
		sumas = "<tr>" &_
		"<td colspan=""" & 12 & """>" &_
			"<center>" &_
						"" &_
			"</center>" &_
		"</td>" &_
								
		celdasumas("SUMAS") &_
		celdasumasnumero("=SUMA(N5:N"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(O5:O"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(P5:P"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Q5:Q"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(U5:U"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(W5:W"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(Y5:Y"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(AA5:AA"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(AC5:AC"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AD5:AD"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AE5:AE"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AF5:AF"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AG5:AG"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AH5:AH"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AI5:AI"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AJ5:AJ"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AK5:AK"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AL5:AL"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AM5:AM"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AN5:AN"&cstr(contador)&")") &_
		celdadatos("")
		celdadatos("")
		sumas =  sumas & "</tr>"
	 	
	   
 	 html = info & header & datos & sumas & "</table><br>" 
		
		
		if Request.Form("Enviar") = "t" then
			set file_FSO = createObject("scripting.filesystemobject")
			if (file_FSO.FileExists("C:\Rep_Imp_Sam.xls")) then
				'response.end()
			end if
			'set Stream = file_FSO.CreateTextFile("C:\Rep_Imp_Sam.xls",true)
			'stream.write(html)
			'stream.close
			if (file_FSO.FileExists("C:\Rep_Imp_Sam.xls")) then
				'a = EnviarEmail("ivan.juarez@rkzego.com","Estimado","ivan.juarez@rkzego.com","Reporte de Impuestos","","",1)
			else
				Response.Write("<strong><br><font color=""#006690"" size=""4"" face=""Arial, Helvetica, sans-serif"">El archivo no pudo se creado.<br> Pongase en contacto con el area de Informàtica.</font></strong>")
			end if
			
			b = ReadExcel("c:\","Rep_Imp_Sam.xls")
					response.write (b)
	response.end()
			'if (file_FSO.FileExists("C:\Rep_Imp_Sam.xls")) then
			'	file_FSO.DeleteFile("C:\Rep_Imp_Sam.xls")
			'end if
			
			
		end if
		
	End If
end if


function ReadExcel(sRuta,sNameFile)
	dim res, Path, ConBD, rsVac, sQuery, sComplemento, count, sQueryFin, ConBD2, rs,comm, FECHVIGOR, FECHPUB
	count= 0
	
	Path = sRuta + sNameFile
	sQuery = "Insert INTO sistemas.enc004repimpsam (t_Pedimento) values "
'-------------------------------------------------------------------------------------------------
	'Leemos el Archivo de Excel
	
	Set ConBD = Server.CreateObject("ADODB.Connection")
	With ConBD
	.Provider = "MSDASQL"
	.ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" & Path & ";ReadOnly=False;"
	.Open
	End With
	Set rsVac = Server.CreateObject("ADODB.Recordset") 

	rsVac.Open "Select * From A4:AQ250 ", ConBD,3,3

	while not rsVac.eof
		if(rsVac.fields.Item("PEDIMENTO").Value <> "" ) then
			sComplemento = sComplemento& " ( '"&rsVac.fields.Item("PEDIMENTO").Value&"') , "  
			count = count + 1
		end if
		rsVac.MoveNext()		
	wend 

	ConBD.Close	
	'-------------------------------------------------------------------------------------------------
	'Insertamos en la BD los registros encontrados
	sQuery = sQuery & sComplemento
	'response.write (sQuery)
	'response.end()
	sQueryFin= mid(sQuery, 1,len(sQuery)-2) &";"

	if( count >=1 )then
		set ConBD2=Server.CreateObject("ADODB.Connection")
			ConBD2.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=sistemas; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

		set comm=Server.CreateObject("ADODB.Command")
		comm.ActiveConnection=ConBD2
		comm.CommandText= sQueryFin
		'Set rs=comm.Execute(ra,parameters,options)
		comm.Execute 
		
		ConBD2.Close
		rs="Se han agregado "&count&" registros con éxito"
	else
		rs="No se encontro ningun registro en el archivo " &sNameFile
	end if
	
	ReadExcel =  rs
	
	
end function

Function EnviarEmail(pFrom, pFromName, pTo, pSubject, pCC, pBCC, pPriority)
strError = ""
	if not pFrom="" and not pTo="" then
			sch = "http://schemas.microsoft.com/cdo/configuration/"
			Set cdoConfig = CreateObject("CDO.Configuration")
			With cdoConfig.Fields
			.Item(sch & "sendusing") = 2
			.Item(sch & "smtpserver") = "smtp.gmail.com"
			.Item(sch & "smtpserverport") = 465
			.Item(sch & "smtpconnectiontimeout") = 30
			.Item(sch & "smtpusessl") = true
			.Item(sch & "smtpauthenticate") = 1
			.Item(sch & "sendusername") = pFrom
			.Item(sch & "sendpassword") = "grk3dEr4"
			.update
			End With

			Set MailObject = Server.CreateObject("CDO.Message")
			Set MailObject.Configuration = cdoConfig
			MailObject.From = pFrom
			MailObject.To = pTo
			MailObject.Subject = pSubject
			MailObject.HTMLBody = "<HTML><p><h3 align=""center""><font face=""Arial, Helvetica, sans-serif"">Ha recibido un correo de la extranet " & cstr(date()) & " A LAS " & cstr(TIME()) & "</font></h3></p><BODY><CENTER>" & pHTML & "</CENTER></BODY></HTML>"
			MailObject.AddAttachment "C:\Rep_Imp_Sam.xls"
		
			On Error Resume Next
			MailObject.Send
			If Err <> 0 Then
				strError = Err.Description
			End If
			
			Set MailObject = Nothing
			Set cdoConfig = Nothing
	end if
EnviarEmail = strError
End Function

function GeneraSQL
	SQL = ""
	condicion = filtro
	SQL = 	"SELECT cast(CONCAT_WS('', i.patent01, '-', i.numped01) as char) Pedimento," &_
			" ifnull(i.refcia01,'-') Referencia, " &_
			" ifnull(fr.fraarn02,'0') FraccionAranc, " &_
			" ifnull(group_concat(fr.d_mer102),'-') Descripcion, " &_
			" i.fecent01 as 'Fecha de Entrada', " &_
			" i.tipcam01 as 'Tipo de Cambio', " &_
			" sum(ifnull(fr.prepag02,0)) 'Valor Merc Mon Nac', " &_
			" format(i.valdol01,2) as 'Valor Dolares', " &_
			" sum(ifnull(fr.cancom02,0)) as 'Total Quantity', " &_
			" sum(ifnull(fr.vmerme02,0)) as 'Invoice Amount', " &_
			" (select ifnull(sum(frac.prepag02),0) from " & strOficina & "_extranet.ssfrac02 as frac where frac.refcia02 = i.refcia01  and frac.patent02 = i.patent01 and frac.adusec02 = i.adusec01 ) as 'Tot Fac Mon Nac', " &_
			" i.valseg01 as 'Valor Seguros', " &_
			" i.segros01 as 'Seguros', " &_
			" i.fletes01 as 'Fletes', " &_
			" i.embala01 as 'Embalajes', " &_
			" i.incble01 as 'OtrosInc', " &_
			" format(sum(ifnull(fr.vaduan02,0)),0) 'Valor Aduana', " &_
			" format(if(i.cveped01 = 'R1',ifnull(cf13.import33,0),ifnull(ifnull(cf1.import36,0),0)),0) DTA, " &_
			" format(sum(ifnull(fr.i_adv102,0) + ifnull(fr.i_adv202,0) + ifnull(fr.i_adv302,0)),0) IGI, " &_
			" format(if(i.cveped01 = 'R1',ifnull(ifnull(cf153.import33,0),0),ifnull(ifnull(cf15.import36,0),0)),0) PRV, " &_
			" format(sum(ifnull(fr.i_iva102,0) + ifnull(fr.i_iva202,0) + ifnull(fr.i_iva302,0)),0) 'IVA', " &_
			" format((ifnull(if(i.cveped01 = 'R1',ifnull(cf13.import33,0),ifnull(cf1.import36,0)),0)  + ifnull(if(i.cveped01 = 'R1',ifnull(cf33.import33,0),ifnull(cf3.import36,0)),0) + ifnull(if(i.cveped01 = 'R1',ifnull(cf63.import33,0),ifnull(cf6.import36,0)),0) + ifnull(if(i.cveped01 = 'R1',ifnull(cf153.import33,0),ifnull(cf15.import36,0)),0) ),0) TotalImpuestos, " &_
			" i.tt_dta01, " &_
			" ifnull(fr.ordfra02,0) as ordfra02," &_
			" ifnull(fr.tasadv02,0) as tasadv02, " &_
			" i.adusec01 as adu, " &_
			" i.patent01 " &_
			
			"from " & strOficina & "_extranet." & tablamov & " as i " &_
			" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
			"      left join " & strOficina & "_extranet.ssfrac02 as fr on fr.refcia02 = i.refcia01  and fr.patent02 = i.patent01 and fr.adusec02 = i.adusec01  " &_
			"           left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1' and cf1.adusec36 = i.adusec01 and cf1.patent36 =i.patent01  " &_
			"            left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3' and cf3.adusec36 = i.adusec01 and cf3.patent36 =i.patent01" &_
			"             left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6' and cf6.adusec36 = i.adusec01 and cf6.patent36 =i.patent01 " &_
			"              left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15' and cf15.adusec36 = i.adusec01 and cf15.patent36 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf13 on cf13.refcia33 = i.refcia01 and cf13.cveimp33 = '1' and cf13.adusec33 = i.adusec01 and cf13.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf33 on cf33.refcia33 = i.refcia01 and cf33.cveimp33 = '3' and cf33.adusec33 = i.adusec01 and cf33.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf63 on cf63.refcia33 = i.refcia01 and cf63.cveimp33 = '6' and cf63.adusec33 = i.adusec01 and cf63.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf153 on cf153.refcia33 = i.refcia01 and cf153.cveimp33 = '15' and cf153.adusec33 = i.adusec01 and cf153.patent33 =i.patent01 " &_
			"where cc.rfccli18 = 'SEM950215S98' " & condicion &_
			" group by i.refcia01, fr.fraarn02,fr.tasadv02  "
				
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
	If IsNull(texto) = True Or texto = "" Then
		texto = "-"
	End If
	cell = 	"<td align=""center"">" &_
				
					texto &_
			
			"</td>"
	celdadatos = cell
end function

function celdanumero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0.00"
	End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumero = cell
end function

function celdanumeroentero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0"
	End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumeroentero = cell
end function


function celdasumas(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "-"
	End If
	cell = 	"<td align=""center"" style=""font-weight: bold"" >" &_
				
					texto &_
	
			"</td>"
	celdasumas = cell
end function

function celdasumasnumero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0.00"
	End If
	cell = 	"<td align=""center"" style=""font-weight: bold"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
	
			"</td>"
	celdasumasnumero = cell
end function

function celdasumasnumeroentero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0"
	End If
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
 valor ="-"
 
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
 valor ="-"
 
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
 valor ="-"
 
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
 valor ="-"
 
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

sqlAct=" select ifnull(cast(group_concat(distinct fact.facmon39) as char),0) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
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
 valor ="-"
 
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

function FechasAct()

sqlAct=" SELECT * FROM registro_monitor WHERE ofic00 in ('RKU','DAI','SAP','CEG','LZR','TOL')"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=intranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()

	Do Until act2.eof
		dato = dato & act2.Fields.Item("ofic00").Value & ":" & act2.Fields.Item("fecha_hora_act").Value & "  || "
	act2.MoveNext()
	Loop
	FechasAct = dato
end function

function conte(referencia,oficina,fraccion,aduana,patente)
dim valor
 valor ="-"
 
if (ucase(oficina) = "ALC")then
	oficina = "LZR"
end if
if (ucase(oficina) = "PAN")then
	oficina = "DAI"
end if

if (ucase(oficina) = "LZR") then
	sqlAct=" select ifnull(group_concat(cont.numcon40),'-') as val from " & oficina & "_extranet.sscont40 as cont  " &_
		" where cont.refcia40 = '" & referencia & "' and cont.patent40 = '" & patente & "' and cont.adusec40 = '" & aduana & "' "
else
	sqlAct=" select ifnull(group_concat(gui1.numgui04),'-') as val from " & oficina & "_extranet.ssguia04 as gui1  " &_
		" where gui1.refcia04 = '" & referencia & "' and gui1.patent04 = '" & patente & "' and gui1.adusec04 = '" & aduana & "' "
end if

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	if not(act2.eof) then
		conte =act2.fields("val").value
	else
		conte =valor
	end if
end function
Function checaCargas

	strSQL = "select count(*) as conteo from intranet.ban_extranet as b where b.m_bandera <> 'NA'"
	
	Set conn = Server.CreateObject ("ADODB.Connection")
	conn.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set recset = CreateObject("ADODB.RecordSet")
	Set recset = conn.Execute(strSQL)
	recset.MoveFirst()
	if recset.Fields.Item("conteo").Value = 0 then
		checaCargas = false
	else
		checaCargas = true
	end if
	
End Function

Function GeneraXML
Dim XML : XML = "<?xml version="1.0"?>" & _ 
"<?mso-application progid="Excel.Sheet"?>" & _ 
"<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"" & _ 
" xmlns:o="urn:schemas-microsoft-com:office:office"" & _ 
" xmlns:x="urn:schemas-microsoft-com:office:excel"" & _ 
" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"" & _ 
" xmlns:html="http://www.w3.org/TR/REC-html40">" & _ 
" <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">" & _ 
"  <Title>::.... REPORTE DE IMPUESTOS SAMSUNG.... ::</Title>" & _ 
"  <Author>Alan Alberto Caballero Ibarra</Author>" & _ 
"  <LastAuthor>Alan Alberto Caballero Ibarra</LastAuthor>" & _ 
"  <Created>" & ISODate & "</Created>" & _ 
"  <Company>Grupo Zego</Company>" & _
"  <Version>14.00</Version>" & _ 
" </DocumentProperties>" & _ 
" <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">" & _ 
"  <AllowPNG/>" & _ 
" </OfficeDocumentSettings>" & _ 
" <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">" & _ 
"  <WindowHeight>9270</WindowHeight>" & _ 
"  <WindowWidth>21315</WindowWidth>" & _ 
"  <WindowTopX>120</WindowTopX>" & _ 
"  <WindowTopY>120</WindowTopY>" & _ 
"  <ProtectStructure>False</ProtectStructure>" & _ 
"  <ProtectWindows>False</ProtectWindows>" & _ 
" </ExcelWorkbook>" & _ 
" <Styles>" & _ 
"  <Style ss:ID="Default" ss:Name="Normal">" & _ 
"   <Alignment ss:Vertical="Bottom"/>" & _ 
"   <Borders/>" & _ 
"   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>" & _ 
"   <Interior/>" & _ 
"   <NumberFormat/>" & _ 
"   <Protection/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="m69142312">" & _ 
"   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s62">" & _ 
"   <Alignment ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s65">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <Font ss:FontName="Arial" x:Family="Swiss" ss:Color="#FFFFFF" ss:Bold="1"/>" & _ 
"   <Interior ss:Color="#006699" ss:Pattern="Solid"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s66">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s67">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <NumberFormat ss:Format="Short Date"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s68">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <NumberFormat ss:Format="#,##0"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s69">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <NumberFormat ss:Format="Standard"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s70">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"" & _ 
"    ss:Bold="1"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s71">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"" & _ 
"    ss:Bold="1"/>" & _ 
"   <NumberFormat ss:Format="Standard"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s72">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"" & _ 
"    ss:Bold="1"/>" & _ 
"   <NumberFormat ss:Format="#,##0"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s73">" & _ 
"   <Borders>" & _ 
"    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"" & _ 
"     ss:Color="#000000"/>" & _ 
"   </Borders>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s77">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="13.5" ss:Color="#000000"" & _ 
"    ss:Bold="1"/>" & _ 
"  </Style>" & _ 
"  <Style ss:ID="s78">" & _ 
"   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom" ss:WrapText="1"/>" & _ 
"   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="12" ss:Color="#000000"" & _ 
"    ss:Bold="1"/>" & _ 
"  </Style>" & _ 
" </Styles>" & _ 
" <Worksheet ss:Name="Rep_Imp_Sam">" & _ 
"  <Table ss:ExpandedColumnCount="42" ss:ExpandedRowCount="14" x:FullColumns="1"" & _ 
"   x:FullRows="1" ss:DefaultColumnWidth="60" ss:DefaultRowHeight="15">" & _ 
"   <Column ss:Width="66.75"/>" & _ 
"   <Column ss:Width="66"/>" & _ 
"   <Column ss:AutoFitWidth="0" ss:Width="55.5"/>" & _ 
"   <Column ss:Width="240" ss:Span="1"/>" & _ 
"   <Column ss:Index="6" ss:Width="183.75"/>" & _ 
"   <Column ss:Width="159.75"/>" & _ 
"   <Column ss:Width="123.75"/>" & _ 
"   <Column ss:Width="63"/>" & _ 
"   <Column ss:Width="102.75"/>" & _ 
"   <Column ss:Width="87"/>" & _ 
"   <Column ss:Width="90.75"/>" & _ 
"   <Column ss:Width="113.25"/>" & _ 
"   <Column ss:Width="150.75"/>" & _ 
"   <Column ss:Width="89.25"/>" & _ 
"   <Column ss:Width="90"/>" & _ 
"   <Column ss:Width="90.75"/>" & _ 
"   <Column ss:Width="153.75"/>" & _ 
"   <Column ss:Width="101.25"/>" & _ 
"   <Column ss:Width="114"/>" & _ 
"   <Column ss:Width="145.5"/>" & _ 
"   <Column ss:Width="75.75"/>" & _ 
"   <Column ss:Width="107.25"/>" & _ 
"   <Column ss:Width="64.5"/>" & _ 
"   <Column ss:Width="96"/>" & _ 
"   <Column ss:Width="87"/>" & _ 
"   <Column ss:Width="118.5"/>" & _ 
"   <Column ss:Width="156"/>" & _ 
"   <Column ss:Width="114"/>" & _ 
"   <Column ss:Width="82.5"/>" & _ 
"   <Column ss:Width="210"/>" & _ 
"   <Column ss:AutoFitWidth="0" ss:Width="24.75"/>" & _ 
"   <Column ss:Width="206.25"/>" & _ 
"   <Column ss:AutoFitWidth="0" ss:Width="24"/>" & _ 
"   <Column ss:Width="177"/>" & _ 
"   <Column ss:AutoFitWidth="0" ss:Width="25.5"/>" & _ 
"   <Column ss:AutoFitWidth="0" ss:Width="24"/>" & _ 
"   <Column ss:Width="205.5"/>" & _ 
"   <Column ss:Width="99"/>" & _ 
"   <Column ss:Width="189.75"/>" & _ 
"   <Column ss:Width="103.5"/>" & _ 
"   <Column ss:AutoFitWidth="0" ss:Width="54.75"/>" & _ 
"   <Row ss:AutoFitHeight="0" ss:Height="17.25">" & _ 
"    <Cell ss:MergeAcross="29" ss:StyleID="s77"><Data ss:Type="String">GRUPO ZEGO</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:AutoFitHeight="0" ss:Height="15.75">" & _ 
"    <Cell ss:MergeAcross="29" ss:StyleID="s78"><Data ss:Type="String">SOLICITUD DE IMPUETOS</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s62"/>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">PEDIMENTO</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">REFERENCIA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">FRACCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">DESCRIPCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">MATERIAL NUMBER</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">FACTURAS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">HOUSE BL o CONTENEDOR(LAZ)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">P/O No</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">INCOTERMS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">FECHA DE ENTRADA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">TIPO DE CAMBIO</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">FACTOR MONEDA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR FACTURA ME</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR MERCANCIA MON NAC</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR DOLARES</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">TOTAL QUANTITY</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">INVOICE AMOUNT</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR TOTAL FACT MON NAC</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">INVOICE CURRENCY</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR SEGUROS PED</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR SEGUROS FRACCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">SEGUROS PED</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">SEGUROS FRACCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">FLETES PED</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">FLETES FRACCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">EMBALAJES PED</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">EMBALAJES FRACCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">OTROS INCREMENTABLES PED</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">OTROS INC FRACCION</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR ADUANA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">VALOR ADUANA CALC</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">DTA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">DTA (CALC)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">IGI</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">IGI (CALC)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">PRV</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">IVA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">IVA (CALC)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">TOTAL IMPUESTOS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">TOTAL IMPUESTOS (CALC)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">CVE TIPO TASA DTA</Data></Cell>" & _ 
"    <Cell ss:StyleID="s65"><Data ss:Type="String">TASA ADV</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="75.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85177001</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">PARTES PARA TELEFONO CELULAR</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">GH59-10949A,GH59-10949A,GH59-10949A,GH59-10653A,GH98-18381C,GH98-18381C,GH98-18744B,GH98-18991B,GH96-04918A,GH98-18377C,GH98-18377C,GH98-18947A,GH98-18947C,GH98-18947C</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5075728744&#5;5076424637&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O5*K5,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1399</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">2121.97</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">34.576000000000001</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N5+W5+Y5+AA5+AC5,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">353</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO5=7,0.008*AD5,SI(AO5=4,af5,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD5* (AP5)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">244</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">5.5759999999999996</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD5+AG5+AI5)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">7.7149999999999999</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG5+AI5+AJ5+AL5,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85229001</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">PARTES PARA REPRODUCTOR DE VIDEO</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">AH96-01311A,AH96-01311A</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5075242624&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O6*K6,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">8</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">47.6</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">776</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N6+W6+Y6+AA6+AC6,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO6=7,0.008*AD6,SI(AO6=4,af6,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD6* (AP6)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">125</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD6+AG6+AI6)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG6+AI6+AJ6+AL6,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85229007</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">CIRCUITOS MODULARES</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">AH94-02843B,AH94-02872A,AH94-02825F</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5075242624&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O7*K7,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">30</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">330.4</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">5.3840000000000003</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N7+W7+Y7+AA7+AC7,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO7=7,0.008*AD7,SI(AO7=4,af7,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD7* (AP7)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">868</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD7+AG7+AI7)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG7+AI7+AJ7+AL7,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85299006</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">CIRCUITOS MODULARES</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">AH92-02739C</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5076424637&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O8*K8,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">4</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">79.08</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1.2889999999999999</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N8+W8+Y8+AA8+AC8,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO8=7,0.008*AD8,SI(AO8=4,af8,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD8* (AP8)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">208</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD8+AG8+AI8)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG8+AI8+AJ8+AL8,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85332101</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">RESISTENCIAS ELECTRICAS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">2007-000143</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5076424637&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O9*K9,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">5</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">0.05</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N9+W9+Y9+AA9+AC9,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO9=7,0.008*AD9,SI(AO9=4,af9,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD9* (AP9)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD9+AG9+AI9)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG9+AI9+AJ9+AL9,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85416001</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">CRISTALES PIEZOELECTRICOS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">2904-001946,2801-005045</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5076424637&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O10*K10,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">90</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">7.4</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">121</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N10+W10+Y10+AA10+AC10,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO10=7,0.008*AD10,SI(AO10=4,af10,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD10* (AP10)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">20</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD10+AG10+AI10)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG10+AI10+AJ10+AL10,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85423199</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">CIRCUITOS INTEGRADOS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">1003-002287</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5076424637&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O11*K11,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">3</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">1.05</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">17</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N11+W11+Y11+AA11+AC11,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO11=7,0.008*AD11,SI(AO11=4,af11,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD11* (AP11)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">3</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD11+AG11+AI11)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG11+AI11+AJ11+AL11,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85444204</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">CONDUCTORES ELECTRICOS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">AH81-07404A,AH81-07405A</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5075242624&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O12*K12,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">10</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">4.1</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">67</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N12+W12+Y12+AA12+AC12,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO12=7,0.008*AD12,SI(AO12=4,af12,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">3</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD12* (AP12)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">11</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD12+AG12+AI12)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG12+AI12+AJ12+AL12,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">5</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="30.75">" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">3945-2002866</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">DAI12-04955</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">85229099</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">ALTAVOCES,PARTES PARA REPRODUCTOR DE VIDEO</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">AH82-00334A,AH82-00334A,AH96-01622A,AH81-05649C</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">9003925362&#5;9003925336&#5;9003921582&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">HKG672820,865 0994 2704</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">5075242624&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">FCA&#5;</Data></Cell>" & _ 
"    <Cell ss:StyleID="s67"><Data ss:Type="DateTime">2012-04-16T00:00:00.000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">130736</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1000000000</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">309.80|1,899.75|497.64</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(O13*K13,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=RC[2]*RC[-3]"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">17</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">115.54</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">35395</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">USD</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-7]/RC[-3])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-9]/RC[-5])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">6968</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-11]/RC[-7])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-13]/RC[-9])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="Number">1751</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69" ss:Formula="=(RC[-15]/RC[-11])*RC[-1]"><Data" & _ 
"      ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">1.883</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(N13+W13+Y13+AA13+AC13,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=SI(AO13=7,0.008*AD13,SI(AO13=4,af13,0))</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR((AD13* (AP13)/100),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">304</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(((AD13+AG13+AI13)*0.16),0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s68"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s69"><Data ss:Type="String">=REDONDEAR(AG13+AI13+AJ13+AL13,0)</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">7</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="Number">0</Data></Cell>" & _ 
"   </Row>" & _ 
"   <Row ss:Height="15.75">" & _ 
"    <Cell ss:MergeAcross="11" ss:StyleID="m69142312"/>" & _ 
"    <Cell ss:StyleID="s70"><Data ss:Type="String">SUMAS</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s72" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">1566</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Error">#VALUE!</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">1025.1320000000001</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">353</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">3</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">244</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">1544.576</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">7.7149999999999999</Data></Cell>" & _ 
"    <Cell ss:StyleID="s71" ss:Formula="=SUM(R[-9]C:R[-1]C)"><Data ss:Type="Number">0</Data></Cell>" & _ 
"    <Cell ss:StyleID="s66"><Data ss:Type="String">-</Data></Cell>" & _ 
"    <Cell ss:StyleID="s73"/>" & _ 
"   </Row>" & _ 
"  </Table>" & _ 
"  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">" & _ 
"   <Print>" & _ 
"    <ValidPrinterInfo/>" & _ 
"    <PaperSizeIndex>9</PaperSizeIndex>" & _ 
"    <HorizontalResolution>600</HorizontalResolution>" & _ 
"    <VerticalResolution>600</VerticalResolution>" & _ 
"   </Print>" & _ 
"   <Selected/>" & _ 
"   <DoNotDisplayGridlines/>" & _ 
"   <Panes>" & _ 
"    <Pane>" & _ 
"     <Number>3</Number>" & _ 
"     <RangeSelection>R1C1:R1C30</RangeSelection>" & _ 
"    </Pane>" & _ 
"   </Panes>" & _ 
"   <ProtectObjects>False</ProtectObjects>" & _ 
"   <ProtectScenarios>False</ProtectScenarios>" & _ 
"  </WorksheetOptions>" & _ 
" </Worksheet>" & _ 
"</Workbook>"
	
End Function

function genera_registros(det,tipope,finicio,ffinal)
dim c,nparte
nparte=""
codigo = ""
 c=chr(34)
 
 codigo=codigo &"<Row>"
 
 genera_html "e","PEDIMENTO","center"
 genera_html "e","REFERENCIA","center"
 genera_html "e","FRACCION","center"
 genera_html "e","DESCRIPCION","center"
 genera_html "e","MATERIAL NUMBER","center"
 genera_html "e","FACTURAS","center"
 genera_html "e","HOUSE BL o CONTENEDOR(LAZ)","center"
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
 genera_html "e","VALOR TOTAL FACT MON NAC","center"
 genera_html "e","INVOICE CURRENCY","center"
 genera_html "e","VALOR SEGUROS PED","center"
 genera_html "e","VALOR SEGUROS FRACCION","center"
 genera_html "e","SEGUROS PED","center"
 genera_html "e","SEGUROS FRACCION","center"
 genera_html "e","FLETES PED","center"
 genera_html "e","FLETES FRACCION","center"
 genera_html "e","EMBALAJES PED","center"
 genera_html "e","EMBALAJES FRACCION","center"
 genera_html "e","OTROS INCREMENTABLES PED","center"
 genera_html "e","OTROS INC FRACCION","center"
 genera_html "e","VALOR ADUANA","center"
 genera_html "e","VALOR ADUANA CALC","center"
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


sqlAct= GeneraSQL

Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
Set RSops = CreateObject("ADODB.RecordSet")
Set RSops = ConnStr.Execute(query)
'response.write(sqlAct)
'rensponse.End()

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 
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
else
	prv = 0
	dta = 0
	totimp = 0
end if

 codigo=codigo &"<Row>"
 genera_html "d",RSops.Fields.Item("Pedimento").Value,"center"
 genera_html "d",RSops.Fields.Item("Referencia").Value,"center"
 genera_html "d",RSops.Fields.Item("FraccionAranc").Value,"center"
 genera_html "d",RSops.Fields.Item("Descripcion").Value,"right"
 genera_html "d",MAN(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d",FACTURAS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d",conte(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d",PO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d",INCOTERMS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d",RSops.Fields.Item("Fecha de Entrada").Value,"right"
 genera_html "d",RSops.Fields.Item("Tipo de Cambio").Value,"right"
 genera_html "d",FACMON(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d",VFME(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value),"right"
 genera_html "d","=REDONDEAR(O" & cstr(contador) & "*K" & cstr(contador) & ",0)","center"
 genera_html "d","=Q" & cstr(contador) & "*L" & cstr(contador),"center"
 genera_html "d",RSops.Fields.Item("Total Quantity").Value,"center"
 genera_html "d",RSops.Fields.Item("Invoice Amount").Value,"right"
 genera_html "d",RSops.Fields.Item("Tot Fac Mon Nac").Value,"right"
 genera_html "d",RSops.Fields.Item("Valor Seguros").Value,"right"
 genera_html "d","=(N" & cstr(contador) & "/R" & cstr(contador) & ")*T" & cstr(contador),"right"
 genera_html "d",RSops.Fields.Item("Seguros").Value,"right"
 genera_html "d","=(N" & cstr(contador) & "/R" & cstr(contador) & ")*V" & cstr(contador),"right"
 genera_html "d",RSops.Fields.Item("Fletes").Value,"right"
 genera_html "d","=(N" & cstr(contador) & "/R" & cstr(contador) & ")*X" & cstr(contador),"right"
 genera_html "d",RSops.Fields.Item("Embalajes").Value,"right"
 genera_html "d","=(N" & cstr(contador) & "/R" & cstr(contador) & ")*Z" & cstr(contador),"center"
 genera_html "d",RSops.Fields.Item("OtrosInc").Value,"center"
 genera_html "d","=(N" & cstr(contador) & "/R" & cstr(contador) & ")*AB" & cstr(contador),"center"
 genera_html "d",RSops.Fields.Item("Valor Aduana").Value,"right"
 genera_html "d","=REDONDEAR(N"&cstr(contador) & "+W"&cstr(contador) & "+Y"&cstr(contador)& "+AA"&cstr(contador) & "+AC"&cstr(contador)& ",0)","right"
 genera_html "d",dta,"center"
 genera_html "d","=SI(AO" & cstr(contador) & "=7,0.008*AD" & cstr(contador) & ",SI(AO" & cstr(contador) & "=4,af" & cstr(contador) & ",0))","center"
 genera_html "d",RSops.Fields.Item("IGI").Value,"center"
 genera_html "d","=REDONDEAR((AD"&cstr(contador) & "* (AP"&cstr(contador) & ")/100),0)","right"
 genera_html "d",prv,"center"
 genera_html "d",RSops.Fields.Item("IVA").Value,"right"
 genera_html "d","=REDONDEAR(((AD"&cstr(contador) & "+AG"&cstr(contador) & "+AI"&cstr(contador) & ")*0.16),0)","center"
 genera_html "d",totimp,"right"
 genera_html "d","=REDONDEAR(AG"&cstr(contador) & "+AI"&cstr(contador) & "+AJ"&cstr(contador)& "+AL"&cstr(contador) & ",0)","center"
 genera_html "d",RSops.Fields.Item("tt_dta01").Value,"right"
 genera_html "d",RSops.Fields.Item("tasadv02").Value,"right"

 codigo=codigo &"</Row>"

refcia =  RSops.Fields.Item("Referencia").Value
Rsops.MoveNext()
Loop
		
		
codigo = Replace((codigo), "á", "a")
codigo = Replace((codigo), "é", "e")
codigo = Replace((codigo), "í", "i")
codigo = Replace((codigo), "ó", "u")
codigo = Replace((codigo), "ú", "u")
'codigo = Replace((codigo), "Á", "A")
'codigo = Replace((codigo), "É", "E")
'codigo = Replace((codigo), "Í", "I")
'codigo = Replace((codigo), "Ó", "O")
'codigo = Replace((codigo), "Ú", "U")
'codigo = Replace((codigo), "´", "")
'codigo = Replace(codigo, ">", ")")
'codigo = Replace(codigo, "<", "(")
'response.Write(codigo)
'response.End()
genera_registros = codigo
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
		<%=html
		%>
	</BODY>
</HTML>

<!--
			Dim ApExcel 
			Set ApExcel = CreateObject("Excel.application")
			ApExcel.Visible = True
			ApExcel.Workbooks.open("C:\Users\alanaci\Desktop\SAMSUNG\Rep_Imp_Sam4.xls")
			ApExcel.Range("AB5:AB8").Select
			ApExcel.Selection.NumberFormat = "#,##0.00"
			celdadatos(RSops.Fields.Item("House B/L").Value) &_
-->
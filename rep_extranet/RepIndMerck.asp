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
	MesIn = UCase(MonthName(Month(fi)))
	DateI = "'" & Anioi & "/" & Mesi & "/" & Diai & "'"

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	MesFi = UCase(MonthName(Month(ff)))
	DateF = "'" & AnioF & "/" & MesF & "/" & DiaF & "'"

Vrfc = Request.Form("rfcCliente")
Vckcve = Request.Form("ckcve")
Vclave = Request.Form("txtCliente")
nivel = request.Form("nivel")
etap = request.Form("etapas")

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
	
	query = ""
	tablamov = ""
	
	if mov = "i" then
		movi = ":: IMPORTACI&Oacute;N ::"
		tablamov = "ssdagi01"
		query = GeneraSQL
	else
		movi = ":: EXPORTACI&Oacute;N ::"
		tablamov = "ssdage01"
		query = GeneraSQL
	end if
	'Response.Write(strOficina)
	'Response.Write(query)
	Set ConnStr = Server.CreateObject("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = Server.CreateObject("ADODB.Recordset")
	Set RSops = ConnStr.Execute(query)
	resultado=0
	resultado=RSops
	
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
		If Tiporepo = "2" Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		RSops.MoveFirst
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
						"<tr>" &_
							"<strong>" &_
								"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
									"<td colspan=""20"">" &_
										"<p align=""left""> <br>" &_
											"<center>" &_
												"<font color=""#000000"" size=""4"" face=""Arial"">" &_
													"<b>" &_
														"REPORTE DE INDICADORES MERCK" &_
													"</b>" &_
												"</font>" &_
											"</center>" &_
										"</p>" &_
										"<p>" &_
										"</p>" &_
										"<p>" &_
										"<center>" &_
												"<font color=""#000000"" size=""4"" face=""Arial"">" &_
													"<b>" &_
														movi &_ 
													"</b>" &_
												"</font>" &_
											"</center>" &_
											
										"</p>" &_
										"<p>" &_
										"</p>" &_
										"<p>" &_
											
											
											"<b>"&_
											"<center>" &_
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
															info = info & "<br>" & "<br>" &"</b>" &_
													"</b>" &_
												"</font>" &_
											"</center>" &_
												
										"</p>" &_
										"<p>" &_
										"</p>" &_
									"</td>" &_
								"</font>" &_
							"</strong>" &_
						"</tr>"
		header = 			"<tr class = ""boton"">" &_
								celdahead("Operaci&oacute;n") &_
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Patente") &_
								celdahead("AduanaSecci&oacute;n") &_
								celdahead("Cliente") &_
								celdahead("Descripci&oacute;nProducto") &_
								celdahead("DTA")
								if mov = "i" Then
									header = header &	celdahead("IGI")
								Else
									header = header & 	celdahead("IGE")
								End If
								header = header & celdahead("IVA") &_
								celdahead("PRV") &_
								celdahead("ECI") &_
								celdahead("FechaPago") &_
								celdahead("FechaEntrada") &_
								celdahead("FechaRevalidaci&oacute;n") &_
								
								celdahead("FechaDespacho")
								if mov = "i" Then
									header = header &	celdahead("KPI FechaDespacho - FechaRevalidaci&oacute;n")
								Else
									header = header & 	celdahead("KPI FechaDespacho - FechaPago")
								End If
								header = header & celdahead("Observaciones KPI") &_
								celdahead("Sem&aacute;foro") &_	
								celdahead("Valor aduana") &_
							"</tr>"
		datos = ""
		
		etapa = ""
		depu=""
		cont = 0
		
		Do Until RSops.Eof
			datos = datos & "<tr>" &_
		
				
								celdadatos(RSops.Fields.Item("operacion").Value) &_
								celdadatos(RSops.Fields.Item("referencia").Value) &_
								celdadatos(RSops.Fields.Item("pedimento").Value) &_
								celdadatos(RSops.Fields.Item("patente").Value) &_
								celdadatos(RSops.Fields.Item("aduanaseccion").Value) &_
								celdadatos(RSops.Fields.Item("cliente").Value) &_
								celdadatos(RSops.Fields.Item("DescripcionProducto").Value) &_
								celdadatos(RSops.Fields.Item("dta").Value)
								
								if mov = "i" Then
									datos = datos &	celdadatos(RSops.Fields.Item("igi").Value)
								Else
									datos = datos &	celdadatos(RSops.Fields.Item("ige").Value)
								End If
				
								datos = datos & celdadatos(RSops.Fields.Item("iva").Value) &_
								celdadatos(RSops.Fields.Item("prv").Value) &_
								celdadatos(RSops.Fields.Item("eci").Value) &_
								
								celdadatos(RSops.Fields.Item("fechapago").Value) &_
								celdadatos(RSops.Fields.Item("fechaEntrada").Value) &_
								celdadatos(RSops.Fields.Item("fechaRevalidacion").Value) &_
								
								celdadatos(RSops.Fields.Item("fechaRobot2").Value) &_
								
								celdadatos(RSops.Fields.Item("kpi").Value) &_
								
								celdadatos(RSops.Fields.Item("kpiEstado").Value) &_
								celdadatos(RSops.Fields.Item("semaforo").Value) &_
								celdadatos(retornaValorAduana(RSops.Fields.Item("referencia").Value,mid(RSops.Fields.Item("referencia").Value,1,3)))
					
			'End If
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

function filtro
	'verifica si esta seleccionado la opcion rfc y le asigna el valor del rfc a condicion para agregarlo a la cadena SQL
	if Vckcve = 0 then
		if IsNumeric(Vrfc) then
			if Vrfc=0 then
				condicion=""
			End if
		else
			condicion = "AND ref.rfccli01 = '" & Vrfc & "' "
		end if
	else
			if Vclave = "Todos" then
				condicion=""
			else
				condicion = "AND ref.cvecli01 = '" & Vclave & "' "
			end if
	end if
	
	filtro = condicion
end function
	
Function GeneraSQL
	condicion=filtro
	SQL = ""
	op=""
	tab=""
	ig=""
	tab2=""
	if mov = "i" Then
		op="impo"
		tab="ssdagi01"
		ig="igi"
		tab2="fecent01"
	else
		op="expo"
		tab="ssdage01"
		ig="ige"
		tab2="fecpre01"
	end if
	
		SQL="select '" & op & "' as operacion,ref.refcia01 as referencia, ref.numped01 as pedimento, ref.patent01 as patente,  " &_
			"ref.adusec01 as aduanaseccion, ref.nomcli01 as cliente,(select group_concat(distinct replace(replace(replace(art.desc05,'\n',''),'\r',''),'\a','')) from " & strOficina & "_extranet.d05artic as art where art.refe05 =ref.refcia01 group by ref.refcia01 ) as DescripcionProducto, " &_
			"if(ref.cveped01='R1',if(dta2.import33 is not null,dta2.import33,0),if(dta.import36 is not null,dta.import36,0)) as 'dta',  " &_
			"if(ref.cveped01='R1',if(" & ig & "2.import33 is not null," & ig & "2.import33,0),if(" & ig & ".import36 is not null," & ig & ".import36,0)) as '" & ig & "', " &_
			"if(ref.cveped01='R1',if(iva2.import33 is not null,iva2.import33,0),if(iva.import36 is not null,iva.import36,0)) as 'iva', " &_
			"if(ref.cveped01='R1',if(prv2.import33 is not null,prv2.import33,0),if(prv.import36 is not null,prv.import36,0)) as 'prv', " &_
			"if(ref.cveped01='R1',if(eci2.import33 is not null,eci2.import33,0),if(eci.import36 is not null,eci.import36,0)) as 'eci', " &_
			"re.feta01 as FechaArrivo, re.fdoc01 as FechaDocumentos, " &_
			"ref.fecpag01 as fechapago, ref." & tab2 & " as fechaEntrada, " &_
			"re.frev01 as fechaRevalidacion," &_
			"if(re.fdoc01>re.feta01,if(re.fdoc01>re.frev01,re.fdoc01,re.frev01),if(re.feta01>re.frev01,re.feta01,re.frev01)) as fechaCompara, " &_
			"ope.Semaforo as semaforo, " &_
			"date(if(MAX(STR_TO_DATE(rob.Fechst01,'%d%m%Y')) is null or MAX(STR_TO_DATE(rob.Fechst01,'%d%m%Y')),if(re.fdsp01='0000-00-00' or re.fdsp01=null,'0000-00-00',re.fdsp01),MAX(STR_TO_DATE(rob.Fechst01,'%d%m%Y')))) AS fechaRobot2,"
			
			if mov="i" then
			
				SQL=SQL & "IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,if(re.frev01='0000-00-00' or re.frev01 is null,ref." & tab2 & ",re.frev01)),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),if(re.frev01='0000-00-00' or re.frev01 is null,ref." & tab2 & ",re.frev01 )))as kpi, " &_
				"if(IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,if(re.frev01='0000-00-00' or re.frev01 is null,ref." & tab2 & ",re.frev01)),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),if(re.frev01='0000-00-00' or re.frev01 is null,ref." & tab2 & ",re.frev01 )))<0 or IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,if(re.frev01='0000-00-00' or re.frev01 is null,ref." & tab2 & ",re.frev01)),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),if(re.frev01='0000-00-00' or re.frev01 is null,ref." & tab2 & ",re.frev01 ))) is null,'Fechas mal capturadas','ok') as kpiEstado " &_
				"from " & strOficina & "_extranet." & tab & " as ref "
			
			else
				SQL=SQL & "IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,ref.fecpag01),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),ref.fecpag01)) as kpi, " &_
				"if(IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,ref.fecpag01),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),ref.fecpag01))<0 or IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,ref.fecpag01),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),ref.fecpag01)) is null,'Fechas mal capturadas','ok') as kpiEstado " &_
				"from " & strOficina & "_extranet." & tab & " as ref "
			end if
			
		SQL=SQL & "left join " & strOficina & "_extranet.c01refer as re on ref.refcia01 =re.refe01  " &_
			"left join " & strOficina & "_extranet.sscont36 as dta on ref.refcia01 = dta.refcia36 and dta.cveimp36 = 1 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as " & ig & " on ref.refcia01 = " & ig & ".refcia36 and " & ig & ".cveimp36 = 6 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as iva on ref.refcia01 = iva.refcia36 and iva.cveimp36 = 3 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as prv on ref.refcia01 = prv.refcia36 and prv.cveimp36 = 15 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as eci on ref.refcia01 = eci.refcia36 and eci.cveimp36 = 18 " &_
			"left join " & strOficina & "_extranet.sscont33 as dta2 on ref.refcia01 = dta2.refcia33 and dta2.cveimp33 = 1 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as " & ig & "2 on ref.refcia01 = " & ig & "2.refcia33 and " & ig & "2.cveimp33 = 6 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as iva2 on ref.refcia01 = iva2.refcia33 and iva2.cveimp33 = 3 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as prv2 on ref.refcia01 = prv2.refcia33 and prv2.cveimp33 = 15 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as eci2 on ref.refcia01 = eci2.refcia33 and eci2.cveimp33 = 18 " &_
			"left join trackingbahia.bit_operaciones as ope on re.refe01 =ope.refcia01  " &_
			"left join trackingbahia.bit_soia as rob on ope.refcia01 =rob.frmsaai01 and rob.Detsit01 in ('710','730') " &_
			"where ref.fecpag01 >=" & DateI & " and ref.fecpag01 <=" & DateF & " and ref.firmae01 is not null and ref.firmae01 != '' and ref.cveped01 != 'R1' " & condicion &_
			"group by ref.refcia01 order by ref.nomcli01 "
'Response.Write(SQL)
'response.end()
	GeneraSQL = SQL
End Function

function celdahead(texto)
	cell = "<td bgcolor = ""#006699"" width=""120"" nowrap>" &_
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

function retornaValorAduana(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="0"
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
  end if
   if (ucase(oficina) = "PAN")then
 oficina = "DAI"
  end if
 
 
sqlAct=" select sum(vaduan02) as campo from "&oficina&"_extranet.ssfrac02 where refcia02 = '"&referencia&"'"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  retornaValorAduana = valor
 else
  retornaValorAduana =valor
 end if


end function
%>

<HTML>
	<HEAD>
		<TITLE>
			:: ....REPORTE DE INDICADORES MERCK.... ::
		</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
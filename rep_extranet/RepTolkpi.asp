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
MesIn = cstr(datepart("m",fi))
AnioI = cstr(datepart("yyyy",fi))
DateI = AnioI & "/" & MesIn & "/" & DiaI

DiaF = cstr(datepart("d",ff))
MesFi = cstr(datepart("m",ff))
AnioF = cstr(datepart("yyyy",ff))
DateF = AnioF & "/" & MesFi & "/" & DiaF

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
	
	Set ConnStr = Server.CreateObject("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = Server.CreateObject("ADODB.Recordset")
	Set RSops = ConnStr.Execute(query)
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
		If Tiporepo = "excel" Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		RSops.MoveFirst
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
						"<tr>" &_
							"<strong>" &_
								"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
									"<td colspan=""22"">" &_
										"<p align=""left"">" &_
											"::.... REPORTE KPI .... ::" &_
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
											"<br>" &_
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
										"</p>" &_
										"<p>" &_
										"</p>" &_
									"</td>" &_
								"</font>" &_
							"</strong>" &_
						"</tr>"
		header = 			"<tr class = ""boton"">" &_
								celdahead("Operacio&oacute;n") &_
								celdahead("Referencia") &_
								celdahead("Pedimento") &_
								celdahead("Patente") &_
								celdahead("AduanaSecci&oacute;n") &_
								celdahead("Cliente") &_
								celdahead("Descripci&oacute;nProducto") &_
								celdahead("DTA") &_
								if mov = "i" Then
									header = header &	celdahead("IGI")
								Else
									header = header & 	celdahead("IGE")
								End If
								celdahead("IVA") &_
								celdahead("PRV") &_
								celdahead("ECI") &_
								celdahead("FechaPago") &_
								celdahead("FechaEntrada") &_
								celdahead("FechaRevalidaci&oacute;n") &_
								celdahead("FechaDespachoRob") &_
								if mov = "i" Then
									header = header &	celdahead("KPI FechaDespacho - FechaPago")
								Else
									header = header & 	celdahead("KPI FechaDespacho - FechaRevalidaci&oacute;n")
								End If
								celdahead("kpiEstado") &_
								celdahead("Sem&aacute;foro") &_	
							"</tr>"
		datos = ""
		referencia = ""
		etapa = ""
		depu=""
		cont = 0
		Set RSmerc = Server.CreateObject("ADODB.Recordset")
		Set RScont = Server.CreateObject("ADODB.Recordset")
		Do Until RSops.Eof
			datos = datos & "<tr>"
			If referencia <> RSops("referencia") Then
				referencia = RSops("referencia")
				
				datos = datos &	celdadatos(referencia) &_
								celdadatos(RSops.Fields.Item("operacion").Value) &_
								celdadatos(RSops.Fields.Item("referencia").Value) &_
								celdadatos(RSops.Fields.Item("pedimento").Value) &_
								celdadatos(RSops.Fields.Item("patente").Value) &_
								celdadatos(RSops.Fields.Item("aduanaseccion").Value) &_
								celdadatos(RSops.Fields.Item("cliente").Value) &_
								celdadatos(RSops.Fields.Item("DescripcinProducto").Value) &_
								celdadatos(RSops.Fields.Item("dta").Value) &_
								
								if mov = "i" Then
									datos = datos &	celdadatos(RSops.Fields.Item("igi").Value)
								Else
									datos = datos &	celdadatos(RSops.Fields.Item("ige").Value)
								End If
				
								celdadatos(RSops.Fields.Item("iva").Value) &_
								celdadatos(RSops.Fields.Item("prv").Value) &_
								celdadatos(RSops.Fields.Item("eci").Value) &_
								
								celdadatos(RSops.Fields.Item("fechapago").Value) &_
								celdadatos(RSops.Fields.Item("fechaEntrada").Value) &_
								celdadatos(RSops.Fields.Item("fechaRevalidacion").Value) &_
								celdadatos(RSops.Fields.Item("fechaRobot2").Value) &_
								
								celdadatos(RSops.Fields.Item("kpi").Value) &_
								
								celdadatos(RSops.Fields.Item("kpiEstado").Value) &_
								celdadatos(RSops.Fields.Item("semaforo").Value)
					
			End If
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
			if Vclave = "Todos"
				condicion=""
			else
				condicion = "AND ref.nomcli01 = '" & Vclave & "' "
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
	if mov = "i" Then
		op="impo"
		tab="ssdagi01"
		ig="igi"
	else
		op="expo"
		tab="ssdage01"
		ig="ige"
	end if
	
		SQL="select '" & op & "' as operacion,ref.refcia01 as referencia, ref.numped01 as pedimento, ref.patent01 as patente,  " &_
			"ref.adusec01 as aduanaseccion, ref.nomcli01 as cliente,(select group_concat(distinct replace(replace(replace(art.desc05,'\n',''),'\r',''),'\a','')) from tol_extranet.d05artic as art where art.refe05 =ref.refcia01 group by ref.refcia01 ) as DescripcionProducto, " &_
			"if(ref.cveped01='R1',if(dta2.import33 is not null,dta2.import33,0),if(dta.import36 is not null,dta.import36,0)) as 'dta',  " &_
			"if(ref.cveped01='R1',if('" & ig & "'2.import33 is not null,'" & ig & "'2.import33,0),if('" & ig & "'.import36 is not null,'" & ig & "'.import36,0)) as ''" & ig & "'', " &_
			"if(ref.cveped01='R1',if(iva2.import33 is not null,iva2.import33,0),if(iva.import36 is not null,iva.import36,0)) as 'iva', " &_
			"if(ref.cveped01='R1',if(prv2.import33 is not null,prv2.import33,0),if(prv.import36 is not null,prv.import36,0)) as 'prv', " &_
			"if(ref.cveped01='R1',if(eci2.import33 is not null,eci2.import33,0),if(eci.import36 is not null,eci.import36,0)) as 'eci', " &_
			"re.feta01 as FechaArrivo, re.fdoc01 as FechaDocumentos, " &_
			"ref.fecpag01 as fechapago, ref.fecent01 as fechaEntrada, " &_
			"re.frev01 as fechaRevalidacion,re.fdsp01 as fechaDespacho, " &_
			"if(re.fdoc01>re.feta01,if(re.fdoc01>re.frev01,re.fdoc01,re.frev01),if(re.feta01>re.frev01,re.feta01,re.frev01)) as fechaCompara, " &_
			"ope.Semaforo as semaforo, " &_
			"max(str_to_date(rob.Fechst01,'%d%m%Y')) as fechaRobot2, "
			
			if mov="i" then
			
				SQL=SQL & "IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,if(re.frev01='0000-00-00' or re.frev01 is null,ref.fecent01,re.frev01)),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),if(re.frev01='0000-00-00' or re.frev01 is null,ref.fecent01,re.frev01 )))as kpi, " &_
				"if(IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,if(re.frev01='0000-00-00' or re.frev01 is null,ref.fecent01,re.frev01)),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),if(re.frev01='0000-00-00' or re.frev01 is null,ref.fecent01,re.frev01 )))<0,"Fechas mal capturadas","ok") as kpiEstado " &_
				"from tol_extranet.'" & tab & "' as ref "
			
			else
				SQL=SQL & "IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,ref.fecpag01),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),ref.fecpag01)) as kpi, " &_
				"if(IF(max(str_to_date(rob.Fechst01,'%d%m%Y')) IS NULL,datediff(re.fdsp01,ref.fecpag01),datediff(max(str_to_date(rob.Fechst01,'%d%m%Y')),ref.fecpag01))<0,"Fechas mal capturadas","ok") as kpiEstado " &_
				"from tol_extranet.'" & tab & "' as ref "
			end if
			
		SQL=SQL & "left join tol_extranet.c01refer as re on ref.refcia01 =re.refe01  " &_
			"left join " & strOficina & "_extranet.sscont36 as dta on ref.refcia01 = dta.refcia36 and dta.cveimp36 = 1 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as '" & ig & "' on ref.refcia01 = '" & ig & "'.refcia36 and '" & ig & "'.cveimp36 = 6 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as iva on ref.refcia01 = iva.refcia36 and iva.cveimp36 = 3 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as prv on ref.refcia01 = prv.refcia36 and prv.cveimp36 = 15 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont36 as eci on ref.refcia01 = eci.refcia36 and eci.cveimp36 = 18 " &_
			"left join " & strOficina & "_extranet.sscont33 as dta2 on ref.refcia01 = dta2.refcia33 and dta2.cveimp33 = 1 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as '" & ig & "'2 on ref.refcia01 = '" & ig & "'2.refcia33 and '" & ig & "'2.cveimp33 = 6 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as iva2 on ref.refcia01 = iva2.refcia33 and iva2.cveimp33 = 3 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as prv2 on ref.refcia01 = prv2.refcia33 and prv2.cveimp33 = 15 " &_
			"LEFT JOIN " & strOficina & "_extranet.sscont33 as eci2 on ref.refcia01 = eci2.refcia33 and eci2.cveimp33 = 18 " &_
			"left join trackingbahia.bit_operaciones as ope on re.refe01 =ope.refcia01  " &_
			"left join trackingbahia.bit_soia as rob on ope.refcia01 =rob.frmsaai01 and rob.Detsit01 in ('710','730') " &_
			"where ref.fecpag01 >='" & DateI & "' and ref.fecpag01 <='" & DateF & "' and ref.firmae01 is not null and ref.firmae01 != '' and ref.cveped01 != 'R1' " & condicion &_
			"group by ref.refcia01 "

	
'	SQL =			"SELECT i.refcia01 AS 'referencia', " &_
'					"CONCAT_WS('-', i.adusec01, i.patent01, i.numped01) AS 'pedimento', " &_
					'"i.cvecli01 AS 'cvecli', " &_
'					"i.nomcli01 AS 'nomcli', " &_
'					"i.totbul01 AS 'bultos', " &_
'					"IF(c.feta01 IS NULL OR c.feta01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.feta01,'%d-%m-%Y')) AS 'feta', " &_
'					"IF(c.fdoc01 IS NULL OR c.fdoc01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.fdoc01,'%d-%m-%Y')) AS 'fdocs', " &_
'					"IF(c.frev01 IS NULL OR c.frev01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.frev01,'%d-%m-%Y')) AS 'frev', " &_
'					"IF(c.fpre01 IS NULL OR c.fpre01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.fpre01,'%d-%m-%Y')) AS 'fprev', " &_
'					"IF(c.fdsp01 IS NULL OR c.fdsp01 = '0000-00-00', '&nbsp;', DATE_FORMAT(c.fdsp01,'%d-%m-%Y')) AS 'fdesp', " &_
'					"IF(i.fecpag01 IS NULL OR i.fecpag01 = '0000-00-00', '&nbsp;', DATE_FORMAT(i.fecpag01,'%d-%m-%Y')) AS 'fpago', "
'	if mov = "i" then
'		SQL = SQL &	"IF(i.fecent01 IS NULL OR i.fecent01 = '0000-00-00' , '&nbsp;', DATE_FORMAT(i.fecent01, '%d-%m-%Y')) as 'fentrada', " &_
'					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) - if( ((DAYOFWEEK(i.fecent01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(i.fecent01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(i.fecent01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) )))),(( TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) - if( ((DAYOFWEEK(i.fecent01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(i.fecent01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(i.fecent01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) * 2) ))))) as 'KPICTE', "
'	else 
'		SQL = SQL & "IF(c.frec01 IS NULL OR c.frec01 = '0000-00-00' , '&nbsp;', DATE_FORMAT(c.frec01, '%d-%m-%Y')) as 'fentrada', " &_
'					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) - if( ((DAYOFWEEK(c.frec01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(c.frec01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(c.frec01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *1.5) )))),(( TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) - if( ((DAYOFWEEK(c.frec01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(c.frec01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(c.frec01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(c.frec01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frec01)) )/ 7)) * 2) ))))) as 'KPICTE', "
'	end if
'	SQL = SQL &		"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) - if( ((DAYOFWEEK(c.frev01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(c.frev01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(c.frev01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *1.5) )))), (( TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) - if( ((DAYOFWEEK(c.frev01) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(c.frev01) -1) = 7 , ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(c.frev01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(c.frev01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.frev01)) )/ 7)) *2) ))))) as 'KPIGRK', " &_
'					"IF(cta.fech31 IS NULL OR cta.fech31 = '0000-00-00','No Hay CG',DATE_FORMAT(MIN(cta.fech31),'%d-%m-%Y')) as 'FCG', " &_
'					"IF(cta.cgas31 IS NULL OR cta.cgas31 = '', 'No Hay CG',MIN(cta.cgas31)) as 'CG', " &_
'					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01) ) - if( ((DAYOFWEEK(c.fdsp01) -1) = 6 ), ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(c.fdsp01) -1) = 7, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(c.fdsp01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01)) )/ 7)) *1.5) )))), (( TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01) ) - if( ((DAYOFWEEK(c.fdsp01) -1) = 6 ), ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *2) - 1, if( (DAYOFWEEK(c.fdsp01) -1) = 7, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01)) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(c.fdsp01) -1)+TO_DAYS(c.fdsp01) - TO_DAYS(c.fdsp01) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(c.fdsp01) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(c.fdsp01)) )/ 7)) * 2) ))))) as 'KPIADMIN', " &_
'					"IF(cta.frec31 IS NULL OR cta.frec31 = '0000-00-00', '&nbsp;', DATE_FORMAT(MIN(cta.frec31), '%d-%m-%Y')) as 'facuse', " &_
'					"IF(MID(i.refcia01,1,3) = 'RKU' OR MID(i.refcia01,1,3) = 'CEG' OR MID(i.refcia01,1,3) = 'SAP' OR MID(i.refcia01,1,3) = 'LZR', (( TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31)) ) - if( ((DAYOFWEEK(MIN(cta.fech31)) -1) = 6 ) , ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *1.5) - 0.5, if( (DAYOFWEEK(MIN(cta.fech31)) -1) = 7 , ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *1.5) - 1, if( ( (DAYOFWEEK(MIN(cta.fech31)) -1)+TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31)) ) = 6, 0.5, ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *1.5) )))), (( TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31)) ) - if( ((DAYOFWEEK(MIN(cta.fech31)) -1) = 6 ), ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *2) - 1, if( (DAYOFWEEK(MIN(cta.fech31)) -1) = 7, ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31))) )/ 7)) *2) - 1, if( ( (DAYOFWEEK(MIN(cta.fech31)) -1)+TO_DAYS(MIN(cta.fech31)) - TO_DAYS(MIN(cta.fech31)) ) = 6, 1, ( FLOOR((( (DAYOFWEEK(MIN(cta.fech31)) -1) + (TO_DAYS(cta.frec31) - TO_DAYS(MIN(cta.fech31))) )/ 7)) * 2) ))))) as 'KPIACUSE', " &_
'					"ets.d_nombre AS 'etapa', " &_
'					"etx.f_fecha AS 'fetapa', " &_
'					"IF(caus.c01causa IS NULL OR caus.c01causa = '', 'Comentarios: ', caus.c01causa) AS 'causa', " &_
'					"IF(etx.m_observ IS NULL OR etx.m_observ = '', '&nbsp;', etx.m_observ) AS 'observacion' " &_
'					"FROM " & strOficina & "_extranet." & tablamov & " AS i " &_
'					"LEFT JOIN " & strOficina & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 " &_
'					"LEFT JOIN " & strOficina & "_extranet.d31refer as ctar ON ctar.refe31 = i.refcia01 " &_
'					"LEFT JOIN " & strOficina & "_extranet.e31cgast as cta ON cta.cgas31 = ctar.cgas31 " &_
'					"LEFT JOIN " & strOficina & "_status.etxpd AS etx ON etx.c_referencia = i.refcia01 " &_ 
'					"LEFT JOIN " & strOficina & "_status.etaps AS ets ON ets.n_etapa = etx.n_etapa " &_
'					"LEFT JOIN " & strOficina & "_status.c01caus AS caus ON caus.c01clavec = etx.clavec AND caus.c01rfc = i.rfccli01 " &_
'					"WHERE i.firmae01 IS NOT NULL AND i.firmae01 <> '' AND i.cveped01 <> 'R1' AND c.fdsp01 >= '" & DateI & "' AND c.fdsp01 <= '" & DateF & "' " &_
'					"AND (cta.esta31 <> 'C' or cta.esta31 IS NULL) AND (cta.fech31 >= c.fdsp01 OR cta.fech31 IS NULL) " &_
'					filtroetapa &_
'					filtrocliente &_
'					"GROUP BY i.refcia01, etx.n_secuenc " &_
'					"ORDER BY i.refcia01, etx.n_etapa "

	' Response.Write(SQL)
	' Response.End()
	GeneraSQL = SQL
End Function

function celdahead(texto)
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
%>

<HTML>
	<HEAD>
		<TITLE>
			:: ....REPORTE DE SEGUIMIENTO DE OPERACIONES.... ::
		</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
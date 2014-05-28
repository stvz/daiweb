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
	
	'if mov = "i" then
	'	movi = ":: IMPORTACI&Oacute;N ::"
	'	tablamov = "ssdagi01"
	'	query = GeneraSQL
	'else
	'	movi = ":: EXPORTACI&Oacute;N ::"
	'	tablamov = "ssdage01"
		query = GeneraSQL
	'end if
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
														"REPORTE DE SERVICIOS COMPLEMENTARIOS" &_
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
								celdahead("Cuenta de gastos") &_
								celdahead("Fecha emisión CG") &_
								celdahead("Esta") &_
								celdahead("Serv Compl CG") &_
								celdahead("Honorarios CG") &_								
								celdahead("Referencia") &_																
								celdahead("Clave ServCompl") &_
								celdahead("Descripción") &_
								celdahead("Importe ") &_
								celdahead("Cliente Facturacion") &_
								celdahead("Aduana") &_
							"</tr>"
		datos = ""
		
		etapa = ""
		depu=""
		cont = 0
		
		Do Until RSops.Eof
			datos = datos & "<tr>" &_				
								celdadatos(RSops.Fields.Item("cgas31").Value) &_
								celdadatos(RSops.Fields.Item("fech31").Value) &_
								celdadatos(RSops.Fields.Item("esta31").Value) &_
								celdadatos(RSops.Fields.Item("csce31").Value) &_
								celdadatos(RSops.Fields.Item("chon31").Value) &_							
								celdadatos(RSops.Fields.Item("refe32").Value) &_								
								celdadatos(RSops.Fields.Item("ttar32").Value) &_
								celdadatos(RSops.Fields.Item("dcrp32").Value) &_
								celdadatos(RSops.Fields.Item("mont32").Value) &_
								celdadatos(RSops.Fields.Item("nomcli18").Value) &_
								celdadatos(jnxadu)
								
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
			condicion = "AND c.rfccli18 = '" & Vrfc & "' "
		end if
	else
			if Vclave = "Todos" then
				condicion=""
			else
				condicion = "AND c.cvecli18 = '" & Vclave & "' "
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
	
	SQL="select e31.cgas31,e31.fech31,e31.esta31,e31.csce31,e31.chon31,  " &_
		"e32.fech32,e32.refe32,e32.totl32,  " &_
		"d32.ttar32,d32.dcrp32,d32.mont32,c.nomcli18  " &_
		"from " & strOficina & "_extranet.e31cgast as e31 " &_
		"inner join " & strOficina & "_extranet.e32rserv as e32 on    e31.cgas31 = e32.cgas32  " &_
		"inner join " & strOficina & "_extranet.ssclie18 as c on c.cvecli18 =e31.clie31   " &_
		"inner join " & strOficina & "_extranet.d32rserv as d32 on d32.refe32  = e32.refe32  " &_
		"where e31.esta31 = 'I' " & condicion & " and e31.fech31 >= "& DateI & " order by fech31,cgas31 "
	
	'Response.Write(SQL)
		
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


%>

<HTML>
	<HEAD>
		<TITLE>
			:: ....REPORTE DE SERVICIOS COMPLEMENTARIOS.... ::
		</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
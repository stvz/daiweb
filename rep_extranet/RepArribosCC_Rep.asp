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
	nocolumns = 7
	tablamov = ""
	if mov = "i" then
		movi = "IMPORTACION"
		tablamov = "ssdagi01"
		query = GeneraSQL(strOficina)
	else
		movi = "EXPORTACION"
		tablamov = "ssdage01"
		query = GeneraSQL(strOficina)
	end if
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	Set RSops = CreateObject("ADODB.RecordSet")
	'response.write(query)
	'response.end()
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
												"DESPACHOS DE " & movi &_
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
								celdahead("Proveedor") &_
								celdahead("Transportista") &_
								celdahead("Fecha de Arribo") &_
								celdahead("Mercancia") &_
								celdahead("Referencia") &_
								celdahead("Observaciones")&_
								celdahead("ODC") 
		header = header &	"</tr>"
		datos = ""
		Referencia = ""
		ubica = ""
		facturas = ""
		contenedores = ""
		total = ""
		importe = 0
		clave=""
		Do Until RSops.EOF
			Referencia = RSops.Fields.Item("referencia").value
					
			obs = Observaciones(Referencia)
			datos = datos &	"<tr>" &_
								celdadatos(RSops.Fields.Item("Proveedor").value) &_
								celdadatos(RSops.Fields.Item("Transportista").Value) &_
								celdadatos(RSops.Fields.Item("FechaArribo").Value) &_
								celdadatos(RSops.Fields.Item("Mercancia").Value) &_
								celdadatos(RSops.Fields.Item("Referencia").Value) &_
								celdadatos(RSops.Fields.Item("Observaciones").Value)&_
								celdadatos(RSops.Fields.Item("ODC").Value)
			datos = datos &	"</tr>"
			Rsops.MoveNext()
		Loop
	
	prom = ""
		html = info & header & datos & "</table><br>" 
	
	End If
end if

Function Observaciones(refe)
	SQLObser = 	""
	observa = ""
	SQLObser = 	"SELECT c_referencia, " &_
				"REPLACE(m_observ,' ','&nbsp;') AS 'obser' " &_
				"FROM " & stroficina & "_status.etxpd " &_
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
		texto = "&nbsp;"
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
		condicion = "AND i.rfccli01 = '" & Vrfc & "' "
	else
		if Vclave <> "Todos" Then
			condicion = "AND i.cvecli01 = " & Vclave & " "
		Else
			permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
			condicion = permi
			condicion = "AND ("&condicion&")"
			if condicion = "AND (cvecli01=0) " Then
				condicion = ""
			end if
		End If
	end if
	filtro = condicion
end function

function GeneraSQL(strADU)
	SQL = ""
	condicion = filtro
	SQL = SQL & "select p.nompro22 Proveedor, t.nom02 Transportista,r.frev01 FechaArribo, f.d_mer102 Mercancia,i.refcia01 Referencia,concat(i.obser101,i.obser102) Observaciones ,d.pedi05 ODC "&_
				"from "&strADU&"_extranet." & tablamov & " as i "&_
				"left join "&strADU&"_extranet.c01refer as r on r.refe01=i.refcia01 "&_
				"left join "&strADU&"_extranet.d01conte as d01 on d01.refe01=i.refcia01 and d01.nemb01<>0 " &_
				"left join "&strADU&"_extranet.e01oemb as e on e.peri01=d01.peri01 and d01.nemb01=e.nemb01 " &_
				"left join "&strADU&"_extranet.c56trans as t on t.cve02=e.ctra01 " &_
				"left join "&strADU&"_extranet.ssfrac02 as f on f.refcia02=i.refcia01 and f.adusec02=i.adusec01 and f.patent02=i.patent01 "&_
				"left join "&strADU&"_extranet.d05artic as d on d.refe05=f.refcia02 and d.agru05=f.ordfra02 and d.frac05=f.fraarn02 "&_
				"left join "&strADU&"_extranet.ssfact39 as fa on fa.refcia39=i.refcia01 and fa.adusec39=i.adusec01 and fa.patent39=i.patent01 and d.fact05=fa.numfac39 "&_
				"left join "&strADU&"_extranet.ssprov22 as p on p.cvepro22=fa.cvepro39 "&_
				" where  i.firmae01 is not null and i.firmae01<>'' and i.cveped01<>'R1' "& condicion&_
				" and r.frev01 between '" & DateI & "' and '" & DateF & "'  "
			

	 'Response.Write(SQL)
	 'Response.End
	GeneraSQL = SQL
end function









%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE ARRIBOS .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
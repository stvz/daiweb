<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000


strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
'permi = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")
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
	end select

	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	Vrfc=Request.Form("rfcCliente")
	filtro = ""
	filtro = "AND i.rfccli01 like '" & Vrfc & "' "
	if fi <> "" then
		DiaI = cstr(datepart("d",fi))
		Mesi = cstr(datepart("m",fi))
		AnioI = cstr(datepart("yyyy",fi))
		DateI = Anioi & "/" & Mesi & "/" & Diai
		filtro = filtro & "AND i.fecpag01 >= '" & DateI & "' "
	Else
		Response.Write("Ingresa una fecha de inicio")
		Response.End()
	End If

	if ff <> "" Then
		DiaF = cstr(datepart("d",ff))
		MesF = cstr(datepart("m",ff))
		AnioF = cstr(datepart("yyyy",ff))
		DateF = AnioF & "/" & MesF & "/" & DiaF
		filtro = filtro & "AND i.fecpag01 <= '" & DateF & "' "
	Else
		Response.Write("Ingresa una fecha final")
		Response.End()
	End If
	
	nocolumns = 8
	query = GeneraSQL
	' Response.Write(query)
	' Response.End()
	
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
		info = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left"">" &_
												"::.... PROVEEDORES .... ::" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
												"Del " & fi & " Al " & ff &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr bgcolor = ""#006699"" class = ""boton"">" &_
								celdahead("Oficina") &_
								celdahead("Cliente") &_
								celdahead("Clave") &_
								celdahead("Proveedor") &_
								celdahead("ID Fiscal") &_
								celdahead("Domicilio") &_
								celdahead("Existe y Afecta") &_
								celdahead("Existe y NO Afecta") &_
								celdahead("No Existe") &_
								celdahead("Referencia-Partida-Vinculacion") &_
							"</tr>"
	datos = ""
	RSops.MoveFirst()
	Do Until RSops.EOF = True
		datos = datos &			"<tr>" &_
									celdadatos(RSops("ofi")) &_
									celdadatos(RSops("nomcli01")) &_
									celdadatos(RSops("cvecli01")) &_
									celdadatos(RSops("nompro39")) &_
									celdadatos(RSops("idfisc39")) &_
									celdadatos(RSops("dompro")) &_
									celdadatos(RSops("Afecta")) &_
									celdadatos(RSops("NoAfecta")) &_
									celdadatos(RSops("Noexiste")) &_
									celdadatos(RSops("Refe-Partida-Vincul")) &_
								"</tr>"
		RSops.MoveNext()
	Loop
	html = info & header & datos & "</table>"
	
	
	End If
end if

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
	If IsNull(cstr(texto)) = True Or cstr(texto) = "" Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"">" &_
				"<font size=""1"" face=""Arial"">" &_
					cstr(texto) &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function GeneraSQL
	SQL = ""
	condicion = filtro
	tabla = ""
	ofi = ""
	conofi = 0
	conops = 0
	For conofi = 1 to 6 Step 1
			select case conofi
				case 1
					strOficina="rku"
				case 2
					strOficina="dai"
				case 3
					strOficina="sap"
				case 4
					strOficina="lzr"
				case 5
					strOficina="ceg"
				case 6
					strOficina="tol"
			end select
		For conops = 1 to 2 step 1
			If conops = 1 Then
				tabla = "ssdagi01"
			Else
				tabla = "ssdage01"
			End If
			
			SQL = SQL &	"SELECT UPPER('" & strOficina & "') AS 'ofi', " &_
						"i.nomcli01, " &_
						"i.cvecli01, " &_
						"fac.nompro39, " &_
						"fac.idfisc39, " &_
						"UPPER(CONCAT_WS(' ', fac.dompro39, fac.noepro39, IF(fac.noipro39 <> '' AND fac.noipro39 IS NOT NULL, CONCAT('Int. ', fac.noipro39), ''), IF(fac.cp_pro39 IS NOT NULL AND fac.cp_pro39 <> '', CONCAT('CP ', fac.cp_pro39), ''), fac.mc_pro39, fac.nomppr39)) AS 'dompro', " &_
						"CONVERT((SELECT GROUP_CONCAT('(', CONCAT_WS('-',a.refcia01, fr.ordfra02, fr.vincul02), ')' SEPARATOR ', ') FROM " & strOficina & "_extranet." & tabla & " AS a INNER JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON a.refcia01 = fr.refcia02 WHERE a.firmae01 <> '' AND a.refcia01 = i.refcia01), CHAR) AS 'Refe-Partida-Vincul', " &_
						"IF(fac.vincul39 = 1, IF((SELECT GROUP_CONCAT('(', CONCAT_WS('-',a.refcia01, fr.ordfra02, fr.vincul02), ')' SEPARATOR ', ') FROM " & strOficina & "_extranet." & tabla & " AS a INNER JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON a.refcia01 = fr.refcia02 AND fr.vincul02 = 2 WHERE a.firmae01 <> '' AND a.refcia01 = i.refcia01) IS NULL, '', 'X'), '') AS 'Afecta', " &_
						"IF(fac.vincul39 = 1, IF((SELECT GROUP_CONCAT('(', CONCAT_WS('-',a.refcia01, fr.ordfra02, fr.vincul02), ')' SEPARATOR ', ') FROM " & strOficina & "_extranet." & tabla & " AS a INNER JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON a.refcia01 = fr.refcia02 AND fr.vincul02 = 1 WHERE a.firmae01 <> '' AND a.refcia01 = i.refcia01) IS NULL, '', 'X'), '') AS 'NOAfecta', " &_
						"IF(fac.vincul39 = 2, 'X', '') AS 'NoExiste' " &_
						"FROM " & strOficina & "_extranet." & tabla & " AS i " &_
						"LEFT JOIN " & strOficina & "_extranet.sspais19 AS pai ON pai.cvepai19 = i.paicli01 " &_
						"LEFT JOIN " & strOficina & "_extranet.ssfact39 AS fac ON i.refcia01 = fac.refcia39 AND i.adusec01 = fac.adusec39 AND i.patent01 = fac.patent39 " &_
						"WHERE i.firmae01 <> '' AND i.firmae01 IS NOT NULL " &_
						"AND (CONCAT(IF(fac.vincul39 = 1, IF((SELECT GROUP_CONCAT('(', CONCAT_WS('-',a.refcia01, fr.ordfra02, fr.vincul02), ')' SEPARATOR ', ') FROM " & strOficina & "_extranet." & tabla & " AS a INNER JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON a.refcia01 = fr.refcia02 AND fr.vincul02 = 2 WHERE a.firmae01 <> '' AND a.refcia01 = i.refcia01) IS NULL, '', 'X'), ''), " &_
						"IF(fac.vincul39 = 1, IF((SELECT GROUP_CONCAT('(', CONCAT_WS('-',a.refcia01, fr.ordfra02, fr.vincul02), ')' SEPARATOR ', ') FROM " & strOficina & "_extranet." & tabla & " AS a INNER JOIN " & strOficina & "_extranet.ssfrac02 AS fr ON a.refcia01 = fr.refcia02 AND fr.vincul02 = 1 WHERE a.firmae01 <> '' AND a.refcia01 = i.refcia01) IS NULL, '', 'X'), ''), " &_
						"IF(fac.vincul39 = 2, 'X', '')) <> '') " &_
						filtro &_
						"GROUP BY i.rfccli01, i.patent01, fac.nompro39, fac.dompro39, fac.vincul39 " 
						If conofi = 6 And conops = 2 Then
							SQL = SQL & "; "
						Else
							SQL = SQL & " UNION "
						End If
		Next
	Next
	' Response.Write(SQL)
	' Response.End
	GeneraSQL = SQL
end function

%>
<HTML>
	<HEAD>
		<TITLE>::.... PROVEEDORES POR CLIENTE .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
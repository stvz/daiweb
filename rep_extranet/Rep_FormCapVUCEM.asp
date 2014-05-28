<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000


strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

'if not permi = "" then
'	permi = "  and (" & permi & ") "
'end if
'AplicaFiltro = False
'strFiltroCliente = ""
'strFiltroCliente = request.Form("txtCliente")


'Tiporepo = Request.Form("TipRep")

'if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
'	blnAplicaFiltro = true
'end if
'if blnAplicaFiltro then
	'permi = " AND cvecli01 =" & strFiltroCliente
'end if
'if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
'	permi = ""
'end if

if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	
	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & Diai

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	nocolumns = 13
	tablamov = ""
	
		
	Dim ConnStr2
	
	Set ConnStr2 = Server.CreateObject ("ADODB.Connection")
	ConnStr2.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=ventanilla_unica; UID=VU; PWD=VUZego; OPTION=16427"
	
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr2.Execute(GeneraSQL)
	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
	
		'if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		'End If
		info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left""><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												"REPORTE DE SELLOS VUCEM.." &_
											"</font></p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p><font color=""#000066"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
												"Del " & fi & " Al " & ff &_
											"</font></p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("Clave") &_
								celdahead("Cliente") &_
								celdahead("RFC") &_
								celdahead("Patente") &_
								celdahead("Tipo de sello") &_
								celdahead("Fecha de registro") &_
								celdahead("Vigencia") &_
								celdahead("Status de revision") &_
								celdahead("Fecha de revision") &_
								celdahead("Observaciones") &_
								celdahead("Correo de Confirmacion") 
				header = header &	"</tr>"
		
		'celdahead("Password del sello") &_
								'celdahead("Clave de web services") &_
		Do Until RSops.EOF

 				datos = datos&"<tr> "&_
								celdadatos(RSops.Fields.Item("Clave").Value)&_
								celdadatos(RSops.Fields.Item("Cliente").Value)&_
								celdadatos(RSops.Fields.Item("RFC Cliente").Value)&_ 
								celdadatos(RSops.Fields.Item("Patente").Value) &_
								celdadatos(RSops.Fields.Item("Tipo de sello").Value) &_
											celdadatos(RSops.Fields.Item("Fecha registro").Value) &_
								celdadatos(RSops.Fields.Item("Vigente").Value) &_
								celdadatos(RSops.Fields.Item("Status de revision").Value) &_
								celdadatos(RSops.Fields.Item("Fecha de revision").Value) &_
								celdadatos(RSops.Fields.Item("Observaciones").Value) &_
								celdadatos(RSops.Fields.Item("Correo").Value) 
						datos = datos &	"</tr>"
							
							'celdadatos(RSops.Fields.Item("Pass sello").Value) &_
								'celdadatos(RSops.Fields.Item("Clave Web Services").Value) &_
			Rsops.MoveNext()
		Loop
	
	Response.Write(info & header & datos & "</table><br>")
	Response.End()
	
'html=info&header& "</table><br>"
	html = info & header & datos & "</table><br>"
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
'On error resume next
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

function GeneraSQL
	SQL = ""
		SQL="SELECT c.i_cve_fiel as 'Clave', c.t_nombre_fiel as 'Cliente',c.t_rfc as 'RFC Cliente', c.n_patent as 'Patente',c.t_sello as 'Tipo de sello',c.t_password_fiel as 'Pass sello'," &_
"c.t_ws_password as 'Clave Web Services',c.f_registro_sello as 'Fecha registro',max(if(c.b_revisado = 3, c.f_vigencia_fiel,null)) as 'Vigente', if(c.b_revisado=1,'Revisado corregir'," &_
" 	if(c.b_revisado=2,'Revisado Aceptado',	 	if(c.b_revisado =3,'Ok',''))) as 'Status de revision',c.f_revision as 'Fecha de revision',c.t_observaciones as 'Observaciones' ," &_
"c.t_correo as 'Correo' FROM ventanilla_unica.cat001_fieles as c" &_
" where c.n_statusborr <> 0 and date(c.f_registro_sello) >='"&DateI&"' and date(c.f_registro_sello) <='"&DateF&"' " &_
"group by c.t_rfc,c.n_patent,c.t_sello ORDER BY c.t_nombre_fiel ,c.f_registro_sello"
		
		
	 
	GeneraSQL = SQL
end function


%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE COMPLETO DE SELLOS VUCEM .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>
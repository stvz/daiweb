<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000000
 

strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli18")

if not permi = "" then
	permi = "  and (" & permi & ") "
end if

if strTipoUsuario = MM_Cod_Admon then
	permi = ""
end if

if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	
	strOficina = ""
	ofi = request.form("ofi")
	Vrfc = Request.Form("rfcCliente")
	Nomcte = Request.Form("txtCliente")
	CvePdto = Request.Form("txtCveProducto")
	Fraccion = Request.Form("txtFraccion")
	Pdto = Request.Form("txtProducto")

	if checaCargas then
		Response.Write("<strong><br><font color=""#006699"" size=""4"" face=""Arial, Helvetica, sans-serif"">Las Bases de Datos se estan actualizando y no es posible llevar a cabo su solicitud. <br> Por Favor intente de nuevo en unos momentos. <br> Gracias.</font></strong>")
		Response.End()
	end if
	
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(GeneraSQL)
	
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
			'inicio reporte
		archive = "Productos_Corporativos"
		
		Response.Write("Cliente") 
		Response.Write(",")
		Response.Write("RFC")
		Response.Write(",")
		Response.Write("Item")
		Response.Write(",")
		Response.Write("Descripción")
		Response.Write(",")
		Response.Write("Observaciones")
		Response.Write(",")
		Response.Write("Fracción") 
		Response.Write(",")
		Response.Write("Oficina")
		Response.Write(",")
		Response.Write("Estatus")
		Response.write vbNewLine
		
		Do Until RSops.EOF
			
				Response.Write(RSops.Fields.Item("nomcli18").Value) 
				Response.Write(",")
				Response.Write(RSops.Fields.Item("rfccli18").Value) 
				Response.Write(",")
				Response.Write(RSops.Fields.Item("item05").Value & "'") 
				Response.Write(",")
				Response.Write(RSops.Fields.Item("desc05").Value) 
				Response.Write(",")
				Response.Write(RSops.Fields.Item("obse05").Value) 
				Response.Write(",")
				Response.Write(RSops.Fields.Item("frac05").Value) 
				Response.Write(",")
				Response.Write(RSops.Fields.Item("oficina").Value)
				Response.Write(",")
				Response.Write(RSops.Fields.Item("status").Value)
				
				Response.write vbNewLine

			Rsops.MoveNext()
		Loop
		   'fin reporte

	Response.AddHeader "Content-Disposition", "attachment; filename=" & archive & "_" & replace(cstr(date()),"/","-") & ".csv;"
	
	End If
end if


function filtro
	
	condicion = " "
	'if (Vrfc <> "0") then
	'	condicion = condicion & "AND rfccli18 in ('" & Vrfc & "') "
	'end if
	if (Nomcte <> "") then
		condicion = condicion & "AND nomcli18 like '%" & Nomcte & "%' "
	end if
	if (CvePdto <> "") then
		condicion = condicion & "AND item05 like '%" & CvePdto & "%' "
	end if
	if (Fraccion <> "") then
		condicion = condicion & "AND trim(frac05) = " & Fraccion & " "
	end if
	if (Pdto <> "") then
		condicion = condicion & "AND desc05 like '%" & Pdto & "%' "
	end if
		
	filtro = condicion
end function


function GeneraSQL
	SQL = ""
	condicion = filtro
	
		if ofi = "a" then
			
			For i = 0 to 5
				
				Select Case i
					Case 0
						strOficina = "RKU"
					Case 1
						strOficina = "DAI"
					Case 2
						strOficina = "TOL"
					Case 3
						strOficina = "SAP"
					Case 4
						strOficina = "LZR"
					Case 5
						strOficina = "CEG"
				End Select
				
				SQL = SQL & " SELECT DISTINCT " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c1i18.nomcli18,'\r',''),'\n',''),char(13),''),',',';') AS nomcli18," & strOficina & "c1i18.rfccli18 AS rfccli18," & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina &   "c05.item05  ,'\r',''),'\n',''),char(13),''),',',';') AS item05, " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c05.item05,'\r',''),'\n',''),char(13),''),',',';') AS item05, " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c05.desc05,'\r',''),'\n',''),char(13),''),',',';') AS desc05, " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c05.obse05,'\r',''),'\n',''),char(13),''),',',';') AS obse05, " & chr(13) & chr(10)
				SQL = SQL & strOficina & "c05.frac05 AS frac05,'" & strOficina & "' AS oficina, " & chr(13) & chr(10)
				SQL = SQL & "IF (" & strOficina & "c05.status05=0,'Sin Estatus',IF(" & strOficina & "c05.status05=1,'Autorizado',IF(" & strOficina & "c05.status05=2,'No Depurado',IF(" & strOficina & "c05.status05=3,'Bloqueado','--')))) AS status  " & chr(13) & chr(10)
				SQL = SQL & " FROM (" & strOficina & "_extranet.c05artic " & strOficina & "c05 join " & strOficina & "_extranet.ssclie18 " & strOficina & "c1i18) " & chr(13) & chr(10)
				SQL = SQL & " where ((" & strOficina & "c05.clie05 = " & strOficina & "c1i18.cvecli18)  " & chr(13) & chr(10)
				SQL = SQL & " and (" & strOficina & "c05.frac05 <> '') and (" & strOficina & "c1i18.rfccli18 <> ''))" & chr(13) & chr(10)
				SQL = SQL & condicion
				
				if (i<>5) then
					SQL = SQL & "UNION ALL " & chr(13) & chr(10)
				else
					'SQL = SQL & "ORDER BY 3,1,2,4 "
				end if
				
			Next
		
		else
		 
			Select Case ofi
				Case "r"
					strOficina = "RKU"
				Case "d"
					strOficina = "DAI"
				Case "t"
					strOficina = "TOL"
				Case "s"
					strOficina = "SAP"
				Case "l"
					strOficina = "LZR"
				Case "c"
					strOficina = "CEG"
			End Select
			 
			 
				SQL = SQL & " SELECT DISTINCT " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c1i18.nomcli18,'\r',''),'\n',''),char(13),''),',',';') AS nomcli18," & strOficina & "c1i18.rfccli18 AS rfccli18," & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina &   "c05.item05  ,'\r',''),'\n',''),char(13),''),',',';') AS item05, " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c05.item05,'\r',''),'\n',''),char(13),''),',',';') AS item05, " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c05.desc05,'\r',''),'\n',''),char(13),''),',',';') AS desc05, " & chr(13) & chr(10)
				SQL = SQL & "replace(replace(replace(replace(" & strOficina & "c05.obse05,'\r',''),'\n',''),char(13),''),',',';') AS obse05, " & chr(13) & chr(10)
				SQL = SQL & strOficina & "c05.frac05 AS frac05,'" & strOficina & "' AS oficina, " & chr(13) & chr(10)
				SQL = SQL & "IF (" & strOficina & "c05.status05=0,'Sin Estatus',IF(" & strOficina & "c05.status05=1,'Autorizado',IF(" & strOficina & "c05.status05=2,'No Depurado',IF(" & strOficina & "c05.status05=3,'Bloqueado','--')))) AS status  " & chr(13) & chr(10)
				SQL = SQL & " FROM (" & strOficina & "_extranet.c05artic " & strOficina & "c05 join " & strOficina & "_extranet.ssclie18 " & strOficina & "c1i18) " & chr(13) & chr(10)
				SQL = SQL & " where ((" & strOficina & "c05.clie05 = " & strOficina & "c1i18.cvecli18)  " & chr(13) & chr(10)
				SQL = SQL & " and (" & strOficina & "c05.frac05 <> '') and (" & strOficina & "c1i18.rfccli18 <> ''))" & chr(13) & chr(10)
				SQL = SQL & condicion

		'fin de query
		end if

	GeneraSQL = SQL
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

%>
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
		case "LZR"
			strOficina="lzr"
	end select

	'cve=request.form("cve")
	'mov=request.form("mov")
	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	Vrfc=Request.Form("rfcCliente")
	Vckcve=Request.Form("ckcve")
	Vclave=Request.Form("txtCliente")
	'response.write(Vrfc & " | ")
	'response.write(Vckcve & " | ")
	'response.write(Vclave & " | ")
	'response.end()
	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & Diai

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	nocolumns = 25
	tablamov = ""
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write("DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	
	 'Response.Write(GeneraSQL)
	' Response.End()
	
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(GeneraSQL)
	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
	
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left""><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												"BOOKLET EXTRA COSTOS COCA-COLA..." &_
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
								celdahead("No.") &_
								celdahead("Referencia") &_
								celdahead("Tipo de operacion") &_
								celdahead("Pedimento") &_
								celdahead("Descripcion de las Mercancias") &_
								celdahead("Contenedores") &_
								celdahead("Total de Bultos") &_
								celdahead("Peso/Kg") &_
								celdahead("Fecha de Despacho") &_
								celdahead("Cuenta de Gastos") &_
								celdahead("Fecha de la C.G.") &_
								celdahead("Honorarios") &_
								celdahead("Servicios Complementarios") &_
								celdahead("Maniobras") &_
								celdahead("Flete Maritimo") &_
								celdahead("Desconsolidacion y revalidacion") &_
								celdahead("Flete Terrestre") &_
								celdahead("Almacenaje") &_
								celdahead("Demoras") &_
								celdahead("Reparacion de contenedor") &_
								celdahead("Retiro de Etiquetas") &_
								celdahead("Estadias de Transporte") &_
								celdahead("Limpieza de contenedor") &_
								celdahead("Reacomodo de contenedores") &_
								celdahead("Total")
				header = header &	"</tr>"
		i=1
		Do Until RSops.EOF

 				datos = datos&"<tr> "&_
								celdadatos(i)&_
								celdadatos(RSops.Fields.Item("Referencia").Value)&_
								celdadatos(RSops.Fields.Item("TipOP").Value)&_
								celdadatos(RSops.Fields.Item("No.Pedimento").Value)&_ 
								celdadatos(RSops.Fields.Item("Cliente").Value) &_
								celdadatos(RSops.Fields.Item("Descripcion Prod.").Value) &_
								celdadatos(RSops.Fields.Item("Contenedores").Value) &_
								celdadatos(RSops.Fields.Item("Bultos").Value) &_
								celdadatos(RSops.Fields.Item("Peso Kg.").Value) &_
								celdadatos(RSops.Fields.Item("F.Desp").Value) &_
								celdadatos(RSops.Fields.Item("Cuenta de gastos").Value) &_
								celdadatos(RSops.Fields.Item("F.Cuenta G").Value) &_
								celdadatos(RSops.Fields.Item("Honorarios").Value) &_
								celdadatos(RSops.Fields.Item("Serv.Complementarios").Value) &_
								celdadatos(RetornaPagosH(strOficina,RSops.Fields.Item("Referencia").Value,RSops.Fields.Item("Cuenta de gastos").Value),"Maniobras") &_
				datos = datos &	"</tr>"
							
			Rsops.MoveNext()
			i+=1
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
	For i = 0 to 3
	
		Select Case i
				Case 0
					aduanaTmp = "rku"
					Oficina="Veracruz"
				Case 1
					aduanaTmp = "dai"
					Oficina="México"
				Case 2
					aduanaTmp = "sap"
					Oficina="Manzanillo"
				Case 3
					aduanaTmp = "lzr"
					Oficina="Lazaro"
						
		End Select
			
				SQL = SQL &	"SELECT '"&Oficina&"' as Oficina, CONCAT_WS(' ', DATE_FORMAT(CG.fech31,'%y'),left(i.adusec01,2), i.patent01, i.numped01) AS 'Pedimento', " & chr(13) & chr(10)
				SQL = SQL & "CG.cgas31 as 'Cuenta de Gastos', " & chr(13) & chr(10)
				SQL = SQL & "i.refcia01 as Referencia, " & chr(13) & chr(10)
				SQL = SQL & "CASE i.cvecli01 " & chr(13) & chr(10)
				SQL = SQL & "WHEN 11004 THEN 'Especiales' " & chr(13) & chr(10)
				SQL = SQL & "WHEN 12000 THEN 'Andina' " & chr(13) & chr(10)
				SQL = SQL & "WHEN 12001 THEN CONCAT_WS(' - ','CAM','Caribe','IC') " & chr(13) & chr(10)
				SQL = SQL & "WHEN 12002 THEN CONCAT_WS(' - ','ConoSur','ExLatam','USA') " & chr(13) & chr(10)
				SQL = SQL & "WHEN 13000 THEN 'Civac - Refacciones y Maquinaria' " & chr(13) & chr(10)
				SQL = SQL & "WHEN 13002 THEN 'Civac - Refacciones y Maquinaria' " & chr(13) & chr(10)
				SQL = SQL & "WHEN 14000 THEN 'Especiales o Proyecto Aeromexico' " & chr(13) & chr(10)
				SQL = SQL & "WHEN 14015 THEN 'Especiales o Proyecto Aeromexico' " & chr(13) & chr(10)
				SQL = SQL & "ELSE '--' " & chr(13) & chr(10)
			SQL = SQL & "END AS Planta, " & chr(13) & chr(10)
			SQL = SQL & "IFNULL((SELECT GROUP_CONCAT(DISTINCT CT.numcon40) FROM "&aduanaTmp&"_extranet.sscont40 CT WHERE CT.refcia40 = i.refcia01 AND CT.patent40 = i.patent01 AND CT.adusec40 = i.adusec01 GROUP BY CT.refcia40),'') as Contenedores, " & chr(13) & chr(10)
			SQL = SQL & "i.nompro01 as 'Nombre de Proveedor', " & chr(13) & chr(10)
			SQL = SQL & "DATE_FORMAT(i.fecpag01,'%d/%m/%Y') as 'Fecha de Pago', " & chr(13) & chr(10)
			SQL = SQL & "REPLACE(CP.desc21,'.','') as Concepto, " & chr(13) & chr(10)
			SQL = SQL & "'' AS 'Dias de Demora o Estadia', " & chr(13) & chr(10)
			SQL = SQL & "sum(If(EP.deha21='A', DP.mont21,(DP.mont21)*-1)) as Monto, " & chr(13) & chr(10)
			SQL = SQL & "'MXN' as Moneda " & chr(13) & chr(10)
			SQL = SQL & "FROM "&aduanaTmp&"_extranet.ssdage01 i " & chr(13) & chr(10)
			SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.c01refer RE ON i.refcia01 = RE.refe01 " & chr(13) & chr(10)
			SQL = SQL & "LEFT JOIN "&aduanaTmp&"_extranet.ssclie18 CC ON cc.cvecli18 = i.cvecli01 " & chr(13) & chr(10)
			SQL = SQL & " INNER JOIN "&aduanaTmp&"_extranet.d31refer RF ON RF.refe31 = i.refcia01 " & chr(13) & chr(10)
			SQL = SQL & "INNER JOIN "&aduanaTmp&"_extranet.e31cgast CG ON CG.cgas31 = RF.cgas31 AND CG.esta31 <> 'C'" & chr(13) & chr(10)
			SQL = SQL & "INNER JOIN "&aduanaTmp&"_extranet.d21paghe DP ON DP.refe21 = i.refcia01 and DP.cgas21 = RF.cgas31 " & chr(13) & chr(10)
			SQL = SQL & "INNER JOIN "&aduanaTmp&"_extranet.e21paghe EP ON EP.foli21=DP.foli21 and year(EP.fech21)=year(DP.fech21) and EP.tmov21=DP.tmov21 and EP.esta21 <> 'S' " & chr(13) & chr(10)
			SQL = SQL & "INNER JOIN "&aduanaTmp&"_extranet.c21paghe CP ON CP.clav21 = EP.conc21 " & chr(13) & chr(10)
				IF Vrfc<>"" then
					if Vrfc="UME651115N48" OR Vrfc="BRM711115GI8" OR Vrfc="ISI011214HM3" OR Vrfc="UMA011214255" OR Vrfc="URE711115AX5" then
						SQL = SQL & "WHERE CC.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3','UMA011214255','URE711115AX5') "& chr(13) & chr(10)
					else
						SQL = SQL & "WHERE CC.rfccli18 ='"&Vrfc&"'"& chr(13) & chr(10)
					end if 
				ELSE
						SQL = SQL & "WHERE CC.cvecli18 ="&Vclave& chr(13) & chr(10)
				END IF
			SQL = SQL & "and i.firmae01 is not null and i.firmae01 <> '' and  CG.fech31 >='"&DateI&"' and CG.fech31 <= '"&DateF&"'"& chr(13) & chr(10)
			if aduanaTmp ="rku" then
			SQL = SQL & "and CP.clav21 not in(31,390,389,406,135,116,25,319,397,1,2) "& chr(13) & chr(10)  
			elseif aduanaTmp="sap" then
			SQL = SQL & "and CP.clav21 not in(188,168,175,17,2,258,70,13,178,12,151,80,127) "& chr(13) & chr(10)  
			elseif aduanaTmp="dai" then
			SQL = SQL &"and CP.clav21 not in(1,2,31,127,239)  "& chr(13) & chr(10) 
			elseif aduanaTmp="lzr" then
			SQL = SQL &"and CP.clav21 not in(1,297)  "& chr(13) & chr(10) 
			end if 
			SQL = SQL & "GROUP BY CG.cgas31,CP.clav21"& chr(13) & chr(10)  
		
		if (i<>3) then
			SQL = SQL & "UNION ALL " & chr(13) & chr(10)
		else
			SQL = SQL & "ORDER BY Referencia DESC "
		end if
	Next
	 'Response.Write(SQL)
	 'Response.End
	GeneraSQL = SQL
end function


%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE EXTRACOSTOS DE COCA-COLA .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>
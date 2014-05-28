<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000000
 

strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli18")
'permi = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")

if not permi = "" then
	permi = "  and (" & permi & ") "
end if

AplicaFiltro = False
strFiltroCliente = ""

 
strFiltroCliente = request.Form("txtCliente")
'strFiltroCliente = request.Form("txtArea")


Tiporepo = Request.Form("TipRep")

if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	blnAplicaFiltro = true
end if
if blnAplicaFiltro then
	permi = " AND cvecli18 =" & strFiltroCliente
	'permi = " AND cvecli18 in" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	permi = ""
end if

if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	
	strOficina = ""
	ff=trim(request.form("ff"))
	ofi=request.form("ofi")
	Vrfc=Request.Form("rfcCliente")
	Vckcve=Request.Form("ckcve")
	Vclave=Request.Form("txtCliente")
	Vcvs=Request.Form("txtArea")
	' response.write(Vrfc & " | ")
	' response.write(Vckcve & " | ")
	' response.write(Vclave & " | ")
	' response.end()

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	
	if checaCargas then
		Response.Write("<strong><br><font color=""#006699"" size=""4"" face=""Arial, Helvetica, sans-serif"">Las Bases de Datos se estan actualizando y no es posible llevar a cabo su solicitud. <br> Por Favor intente de nuevo en unos momentos. <br> Gracias.</font></strong>")
		Response.End()
	end if
	
	if (Vcvs = "") then
		if (Vckcve = 0) Then
			if Vrfc = "0" and Tiporepo = 1 then
				Response.Write("No puede generar el reporte detallado de todos los clientes.")
				Response.End()
			end if
		else
			if Vclave = "Todos" and Tiporepo = 1 then
				Response.Write("No puede generar el reporte detallado de todos los clientes.")
				Response.End()
			end if
		end if
	else
		if Vcvs = "0" and Tiporepo = 1 then
				Response.Write("No puede generar el reporte detallado de todos los clientes.")
				Response.End()
		else
			if Vcvs = "Todos" and Tiporepo = 1 then
				Response.Write("No puede generar el reporte detallado de todos los clientes.")
				Response.End()
			end if
		end if
	end if

	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Response.Write("DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=" & strOficina & "_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
	' Response.Write(query & "<br><br>")
	' Response.Write(Actualizaciones)
	
	' Response.Write(GeneraSQL)
	' Response.End()
	
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(GeneraSQL)
	IF RSops.BOF = True And RSops.EOF = True Then
		Response.Write("No hay datos para esas condiciones")
	Else
		
		'Response.Addheader "Content-Disposition", "attachment;"
		'Response.ContentType = "application/vnd.ms-excel"
		
		select case Tiporepo 
		
		case 2
			
			archive = "RepCarteraGeneral"
			Response.Write("REPORTE DE CARTERA GENERAL")
			Response.write vbNewLine
			Response.Write("Fecha de Corte Al " & ff)
			Response.write vbNewLine
		
			Response.Write("Clave del Cliente")
			Response.Write(",")
			Response.Write("RFC del Cliente")
			Response.Write(",")
			Response.Write("Nombre del Cliente")
			Response.Write(",")
			Response.Write("Oficina")
			Response.Write(",")
			Response.Write("Saldo Vencido ECG")
			Response.Write(",")
			Response.Write("Saldo Vigente ECG")
			Response.Write(",")
			Response.Write("Saldo Vencido RCG")
			Response.Write(",")
			Response.Write("Saldo Vigente RCG")
			Response.Write(",")
			Response.Write("Saldo a Favor")
			Response.Write(",")
			Response.Write("Saldo NA")
			Response.Write(",")
			Response.Write("Saldo Total")
			Response.Write(",")
			Response.Write("Dias de Credito")
			Response.Write(",")
			Response.Write("Financiamiento")
			Response.Write(",")
			Response.Write("Cantidad de CG")
			Response.Write(",")
			Response.Write("Fecha Ultimo Movimiento")
			Response.Write(",")
			' Response.Write("Pagos Hechos del Saldo")
			' Response.Write(",")
			' Response.Write("Honorarios del Saldo")
			' Response.Write(",")
			' Response.Write("Servicios Complementarios del Saldo")
			' Response.Write(",")
			' Response.Write("IVA Trasladado del Saldo")
			' Response.Write(",")
			' Response.Write("Anticipos Sin Aplicar")
			' Response.Write(",")
			' Response.Write("Pagos Hechos Sin Aplicar")
			' Response.Write(",")
			Response.Write("Ingresos Promedio Por Mes")
			Response.Write(",")
			Response.Write("Operaciones Promedio Por Mes")
			Response.Write(",")
			Response.Write("Cantidad de Meses")
			Response.Write(",")
			Response.Write("Cantidad de Operaciones por Mes")
			Response.write vbNewLine
		
			'celdahead("Proveedor") &_
			''celdahead("Descripcion de la Mercancia") &_
									
			
			Do Until RSops.EOF
 
				
					Response.Write(RSops.Fields.Item("Clave del Cliente").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("RFC del Cliente").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Nombre del Cliente").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Oficina").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo Vencido ECG").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo Vigente ECG").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo Vencido RCG").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo Vigente RCG").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo a Favor").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo NA").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("SaldoTotal").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Dias de Credito").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Financiamiento").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Cantidad de CG").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Fecha Ultimo Movimiento").Value)
					Response.Write(",")
					'Response.Write(RSops.Fields.Item("Pagos Hechos del Saldo").Value)
					'Response.Write(",")
					'Response.Write(RSops.Fields.Item("Honorarios del Saldo").Value)
					'Response.Write(",")
					'Response.Write(RSops.Fields.Item("Servicios Complementarios del Saldo").Value)
					'Response.Write(",")
					'Response.Write(RSops.Fields.Item("IVA Trasladado del Saldo").Value)
					'Response.Write(",")
					'Response.Write(RSops.Fields.Item("Anticipos Sin Aplicar").Value)
					'Response.Write(",")
					'Response.Write(RSops.Fields.Item("Pagos Hechos Sin Aplicar").Value)
					'Response.Write(",")
					Response.Write(RSops.Fields.Item("Ingresos Promedio Por Mes").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Operaciones Promedio Por Mes").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Cantidad de Meses").Value)
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Cantidad de Operaciones por Mes").Value)
					Response.write vbNewLine
				
								
				Rsops.MoveNext()
			Loop
			
		case 1 'Detallado
			
			archive = "RepCarteraDetallado"
			Response.Write("REPORTE DE CARTERA DETALLADO")
			Response.write vbNewLine
			Response.Write("Fecha de Corte Al " & ff)
			Response.write vbNewLine

			Response.Write("Fecha del Movimiento") 
			Response.Write(",")
			Response.Write("Asiento") 
			Response.Write(",")
			Response.Write("Cuenta de Gastos") 
			Response.Write(",")
			Response.Write("Concepto") 
			Response.Write(",")
			Response.Write("Usuario Captura") 
			Response.Write(",")
			Response.Write("Folio") 
			Response.Write(",")
			Response.Write("Oficina") 
			Response.Write(",")
			Response.Write("Referencia") 
			Response.Write(",")
			Response.Write("RFC") 
			Response.Write(",")
			Response.Write("Cliente") 
			Response.Write(",")
			Response.Write("Clave de Cliente") 
			Response.Write(",")
			Response.Write("Plazo de Credito") 
			Response.Write(",")
			Response.Write("Cuenta Contable") 
			Response.Write(",")
			Response.Write("Tipo de Financiamiento") 
			Response.Write(",")
			Response.Write("Fecha de Recepcion de CG") 
			Response.Write(",")
			Response.Write("Saldo") 
			Response.Write(",")
			Response.Write("Status de Cuenta de Gastos") 
			Response.Write(",")
			Response.Write("Dias Pendientes") 
			Response.Write(",")
			Response.Write("Tipo de Cartera (Recepcion de la Cuenta de Gastos)") 
			Response.Write(",")
			Response.Write("Tipo de Cartera (Emision de la Cuenta de Gastos)")
			Response.write vbNewLine
		
			'celdahead("Proveedor") &_
			''celdahead("Descripcion de la Mercancia") &_
									

			Do Until RSops.EOF
				
					Response.Write(RSops.Fields.Item("fech11").Value) 
					Response.Write(",")
					Response.Write("'" & RSops.Fields.Item("Asie11").Value & "'") 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("refe11").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("conc11").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("user11").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("d11_folres11").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("facofna").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("refe01").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("cvecli18").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("rfccli18").Value) 
					Response.Write(",")
					Response.Write(replace(RSops.Fields.Item("nomcli18").Value,","," ")) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("plcred18").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("CuentaContable").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("financ18").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("frec31").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("statusCG").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("DiasPendientes").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("TipoCarteraRCG").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("TipoCartera").Value)
					Response.write vbNewLine
								
				Rsops.MoveNext()
			Loop
			
		case 3
		
			'inicio reporte
			archive = "RepCarteraEdoCta"
			Response.Write("REPORTE DE ESTADO DE CUENTA")
			Response.write vbNewLine
			Response.Write("Fecha de Corte Al " & ff)
			Response.write vbNewLine
			
			Response.Write("Aduana") 
			Response.Write(",")
			Response.Write("Clave del Cliente") 
			Response.Write(",")
			Response.Write("Nombre del Cliente") 
			Response.Write(",")
			Response.Write("Fecha del Movimiento") 
			Response.Write(",")
			Response.Write("Fecha de Recepcion de CG") 
			Response.Write(",")
			Response.Write("Cuenta de Gastos") 
			Response.Write(",")
			Response.Write("Anticipos") 
			Response.Write(",")
			Response.Write("Saldo Vencido RCG") 
			Response.Write(",")
			Response.Write("Saldo Vigente RCG") 
			Response.Write(",")
			Response.Write("Saldo a Favor") 
			Response.Write(",")
			Response.Write("Saldo NA") 
			Response.Write(",")
			Response.Write("Saldo Total") 
			Response.write vbNewLine
		
			'celdahead("Proveedor") &_
			''celdahead("Descripcion de la Mercancia") &_
									

			Do Until RSops.EOF
				
					Response.Write(RSops.Fields.Item("adu").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("cvecli18").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("nomcli18").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("fech11").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("frec31").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("refe11").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Anticipos").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo Vencido RCG").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo Vigente RCG").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo a Favor").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("Saldo NA").Value) 
					Response.Write(",")
					Response.Write(RSops.Fields.Item("SaldoTotal").Value) 
			    	Response.write vbNewLine
								
				Rsops.MoveNext()
			Loop
			   'fin reporte
			   
		End select

	' Response.Write(info & header & datos & "</table><br>" & prom)
	' Response.End()
	
	Response.AddHeader "Content-Disposition", "attachment; filename=" & archive & "_" & replace(cstr(date()),"/","-") & "_" & replace(replace(replace(cstr(time()),":","")," p.m.","")," a.m.","") & ".csv;"
	
	End If
end if


function celdahead(texto)
	cell = "<td bgcolor = ""#1B1B79"" align=""center"">" & _
					"<strong>" & _
						"<font color=""#FFFFFF"" size=""2"" face=""Calibri,Arial, Helvetica, sans-serif"">" & _
							texto & _
						"</font>" & _
					"</strong>" & _
			"</td>"
	celdahead = cell
end function

function celdadatos(texto)
'On error resume next
	If IsNull(texto) = True Or texto = "" Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"" nowrap>" &_
				"<font size=""2"" face=""Calibri,Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function filtro
	if (Vcvs = "") then
		if Vckcve = 0 then
			if Vrfc <> "0" then
				condicion = "AND rfccli18 = '" & Vrfc & "' "
			else
				'condicion = "AND i.rfccli01 = 'UME651115N48' "
				condicion = " "
			end if
		else
			if Vclave <> "Todos" Then
				condicion = "AND ccli11 = " & Vclave & " "
			Else
				permi = PermisoClientes(Session("GAduana"),strPermisos,"ccli11")
				condicion = permi
				condicion = "AND " & condicion
				if condicion = "AND ccli11=0 " Then
					condicion = ""
				end if
			End If
		end if
	else	
		if Vckcve = 0 then
			if Vcvs <> "0" then
				condicion = "AND rfccli18 in ('" & Vcvs & "') "
			else
				'condicion = "AND i.rfccli01 = 'UME651115N48' "
				condicion = " "
			end if
		else
			if Vclave <> "Todos" Then
				condicion = "AND ccli11 in (" & Vcvs & ") "
			Else
				permi = PermisoClientes(Session("GAduana"),strPermisos,"ccli11")
				condicion = permi
				condicion = "AND " & condicion
				if condicion = "AND ccli11=0 " Then
					condicion = ""
				end if
			End If
		end if
	end if
	
	filtro = condicion
end function


function GeneraSQL
	SQL = ""
	condicion = filtro
	
	select case Tiporepo
	
	case 2 'REPORTE GENERAL
	
		if ofi = "a" then
				
				For i = 0 to 5
					
					Select Case i
							Case 0
								strOficina = "rku"
								adu = "VERACRUZ"
								facofna = "0001"
							Case 1
								strOficina = "dai"
								adu = "MEXICO"
								facofna = "0005"
							Case 2
								strOficina = "tol"
								adu = "TOLUCA"
								facofna = "0010"
							Case 3
								strOficina = "sap"
								adu = "MANZANILLO"
								facofna = "0004"
							Case 4
								strOficina = "lzr"
								adu = "LAZARO CARDENAS"
								facofna = "0009"
							Case 5
								strOficina = "ceg"
								adu = "ALTAMIRA"
								facofna = "0003"
					End Select
					
					SQL = SQL & "SELECT bc.cvecli18 as 'Clave del Cliente',bc.rfccli18 'RFC del Cliente',replace(bc.nomcli18,',',' ') as 'Nombre del Cliente','" & adu & "' AS Oficina, " & chr(13) & chr(10)
					SQL = SQL & "MAX(bc.DiasCredito) AS 'Dias de Credito',IF(bc.financ18 = 2, 'Si',IF(bc.financ18 = 1, 'No', 'S/CAP') ) AS Financiamiento, " & chr(13) & chr(10)

					SQL = SQL & "COUNT( IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA')  AND bc.asie11 = '1109000000000000',bc.refe11,NULL) ) AS 'Cantidad de CG', " & chr(13) & chr(10)
					SQL = SQL & "DATE_FORMAT(MAX(bc.fech11),'%d/%m/%Y') AS 'Fecha Ultimo Movimiento', " & chr(13) & chr(10)

					SQL = SQL & "IF( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) + " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) <0,0, " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) + " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
					SQL = SQL & ")  AS 'Pagos Hechos del Saldo', " & chr(13) & chr(10)

					SQL = SQL & "IF(SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
					SQL = SQL & "<0, " & chr(13) & chr(10)
					SQL = SQL & "( ( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000100000000',bc.Saldo ,0))   / SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('4100000100000000', '4100000200080000','2111000000000000','2106000000000000'),bc.Saldo,0))  )  * (  SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) )  )  + SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000100000000',bc.Saldo ,0))  , " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000100000000',bc.Saldo ,0)) ) " & chr(13) & chr(10)
					SQL = SQL & "AS 'Honorarios del Saldo', " & chr(13) & chr(10)

					SQL = SQL & "IF(SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
					SQL = SQL & "<0, " & chr(13) & chr(10)
					SQL = SQL & "( ( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000200080000',bc.Saldo ,0))   / SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('4100000100000000', '4100000200080000','2111000000000000','2106000000000000'),bc.Saldo,0))  )  * (  SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) )  )  + SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000200080000',bc.Saldo ,0))  , " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000200080000',bc.Saldo ,0)) ) " & chr(13) & chr(10)
					SQL = SQL & "AS 'Servicios Complementarios del Saldo', " & chr(13) & chr(10)

					SQL = SQL & "IF(SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
					SQL = SQL & "<0, " & chr(13) & chr(10)
					SQL = SQL & "( ( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('2111000000000000','2106000000000000'),bc.Saldo ,0))   / SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('4100000100000000', '4100000200080000','2111000000000000','2106000000000000'),bc.Saldo,0))  )  * (  SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') " & chr(13) & chr(10)

					SQL = SQL & "AND bc.asie11 IN ('1103000100000000') " & chr(13) & chr(10)

					SQL = SQL & ", bc.Saldo,0) ) )  )  + SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('2111000000000000','2106000000000000'),bc.Saldo ,0))  , " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('2111000000000000','2106000000000000'),bc.Saldo ,0)) ) " & chr(13) & chr(10)
					SQL = SQL & "AS 'IVA Trasladado del Saldo', " & chr(13) & chr(10)

					SQL = SQL & "SUM(IF(bc.TipoCartera = 'CARTERA VENCIDA',bc.Saldo,0)) AS 'Saldo Vencido ECG', " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.TipoCartera = 'CARTERA NORMAL',bc.Saldo,0)) AS 'Saldo Vigente ECG', " & chr(13) & chr(10)

					SQL = SQL & "SUM(IF(bc.TipoCarteraRCG = 'CARTERA VENCIDA',bc.Saldo,0)) AS 'Saldo Vencido RCG', " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.TipoCarteraRCG = 'CARTERA NORMAL',bc.Saldo,0)) AS 'Saldo Vigente RCG', " & chr(13) & chr(10)

					SQL = SQL & "SUM(IF(bc.TipoCartera = 'LIQUIDADO',bc.Saldo,0)) AS 'Saldo a Favor', " & chr(13) & chr(10)
					SQL = SQL & "SUM(IF(bc.TipoCartera = 'N/A',bc.Saldo,0)) AS 'Saldo NA', " & chr(13) & chr(10)
					SQL = SQL & "SUM(bc.Saldo) AS SaldoTotal, " & chr(13) & chr(10)

					SQL = SQL & "IFNULL((SELECT SUM(   IF(bc1.conc11 IN ('ANT','CF2'),bc1.mont11, IF(bc1.conc11 IN ('FA2','CAN'),bc1.mont11*-1,0) ) ) " & chr(13) & chr(10)
					SQL = SQL & "FROM trackingbahia.Bit_Cartera_" & strOficina & " AS bc1 WHERE bc1.conc11 IN ('ANT','FA2','CF2','CAN')  AND bc1.cont11 = 'A'  AND bc1.cvecli18=bc.cvecli18  AND bc1.refe01 = bc.refe01 GROUP BY bc1.cvecli18 ),0) AS 'Anticipos Sin Aplicar', " & chr(13) & chr(10)

					SQL = SQL & "IFNULL((SELECT SUM( phe.mont21) " & chr(13) & chr(10)
					SQL = SQL & "FROM " & strOficina & "_extranet.d21paghe AS phe " & chr(13) & chr(10)
					SQL = SQL & "INNER JOIN " & strOficina & "_extranet.e21paghe AS e21 ON e21.foli21 = phe.foli21 AND YEAR(e21.fech21) = YEAR(phe.fech21) AND e21.tmov21 = phe.tmov21 " & chr(13) & chr(10)
					SQL = SQL & "INNER JOIN " & strOficina & "_extranet.c01refer AS r21 ON r21.refe01 = phe.refe21 " & chr(13) & chr(10)
					SQL = SQL & "WHERE phe.esta21 <> 'S' AND " & chr(13) & chr(10)
					SQL = SQL & "e21.esta21 NOT IN ('S','A') AND " & chr(13) & chr(10)
					SQL = SQL & "r21.clie01 = bc.cvecli18 AND phe.cgas21 =  '' AND " & chr(13) & chr(10)
					SQL = SQL & "r21.refe01 NOT IN (SELECT DISTINCT d31.refe31 FROM " & strOficina & "_extranet.d31refer AS d31 INNER JOIN " & strOficina & "_extranet.c01refer AS rr ON rr.refe01 = d31.refe31 WHERE rr.clie01 = r21.clie01) " & chr(13) & chr(10)
					SQL = SQL & "GROUP BY r21.clie01),0) AS 'Pagos Hechos Sin Aplicar', " & chr(13) & chr(10)

					SQL = SQL & "IFNULL(( ROUND(IFNULL((    (SUM(IF( (bc.facofna = '" & facofna & "'  OR bc.facofna = '')  AND bc.ASIE11 = '4100000100000000', IF(bc.CONC11 REGEXP 'FA1|SCA|DEV|CAR|FA2' , bc.MONT11, IF(bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|CF2' , bc.MONT11*-1,0)), 0)) " & chr(13) & chr(10)
					SQL = SQL & "+ SUM(IF( (bc.facofna = '" & facofna & "'  OR bc.facofna = '')  AND bc.ASIE11 = '4100000200080000',IF(bc.CONC11 REGEXP 'FA1|SCA|DEV|CAR|FA2' , bc.MONT11, IF(bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|CF2' , bc.MONT11*-1, 0)),0))  ) " & chr(13) & chr(10)
					SQL = SQL & "),0),2) /  IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fech11),LPAD(MONTH(rt.fech11),2,'0')  )) )  FROM trackingbahia.Bit_Cartera_" & strOficina & " AS rt WHERE rt.cvecli18 = bc.cvecli18 GROUP BY rt.cvecli18),0) ),0) AS 'Ingresos Promedio Por Mes', " & chr(13) & chr(10)

					SQL = SQL & "IFNULL(( ( IFNULL(( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdagi01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND  ri.firmae01 <> '' AND ri.cveped01 <> 'R1'),0) + " & chr(13) & chr(10)
					SQL = SQL & "IFNULL(( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdage01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND  ri.firmae01 <> '' AND ri.cveped01 <> 'R1'),0) ) / " & chr(13) & chr(10)
					SQL = SQL & "IF( IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) > " & chr(13) & chr(10)
					SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
					SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
					SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) " & chr(13) & chr(10)
					SQL = SQL & ") " & chr(13) & chr(10)
					SQL = SQL & "),0) " & chr(13) & chr(10)
					SQL = SQL & "AS 'Operaciones Promedio Por Mes', " & chr(13) & chr(10)

					SQL = SQL & "( IF( IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) > " & chr(13) & chr(10)
					SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
					SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
					SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) " & chr(13) & chr(10)
					SQL = SQL & "))  AS 'Cantidad de Meses', " & chr(13) & chr(10)

					SQL = SQL & "( ( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdagi01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND ri.firmae01 <> '' AND ri.cveped01 <> 'R1') + " & chr(13) & chr(10)
					SQL = SQL & "( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdage01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND ri.firmae01 <> '' AND ri.cveped01 <> 'R1') )  AS 'Cantidad de Operaciones por Mes' " & chr(13) & chr(10)

					SQL = SQL & "FROM trackingbahia.Bit_Cartera_" & strOficina & "  AS bc " & chr(13) & chr(10)
					SQL = SQL & "WHERE " & chr(13) & chr(10)
					SQL = SQL & "bc.cont11 <> 'C' AND " & chr(13) & chr(10)
					SQL = SQL & "bc.fech11 <= '" & DateF & "' " & chr(13) & chr(10)
					SQL = SQL & condicion & chr(13) & chr(10)
					SQL = SQL & "GROUP BY bc.ccli11 " & chr(13) & chr(10)
					SQL = SQL & "HAVING SaldoTotal NOT BETWEEN -1 AND 0.1 " & chr(13) & chr(10)
					
					if (i<>5) then
						SQL = SQL & "UNION ALL " & chr(13) & chr(10)
					else
						SQL = SQL & "ORDER BY 2,4,1 ASC "
					end if
					
				Next
				
		else
				Select Case ofi
						Case "r"
							strOficina = "rku"
							adu = "VERACRUZ"
							facofna = "0001"
						Case "d"
							strOficina = "dai"
							adu = "MEXICO"
							facofna = "0005"
						Case "t"
							strOficina = "tol"
							adu = "TOLUCA"
							facofna = "0010"
						Case "s"
							strOficina = "sap"
							adu = "MANZANILLO"
							facofna = "0004"
						Case "l"
							strOficina = "lzr"
							adu = "LAZARO CARDENAS"
							facofna = "0009"
						Case "c"
							strOficina = "ceg"
							adu = "ALTAMIRA"
							facofna = "0003"
				End Select

				SQL = SQL & "SELECT bc.cvecli18 as 'Clave del Cliente',bc.rfccli18 'RFC del Cliente',replace(bc.nomcli18,',',' ') as 'Nombre del Cliente','" & adu & "' AS Oficina, " & chr(13) & chr(10)
				SQL = SQL & "MAX(bc.DiasCredito) AS 'Dias de Credito',IF(bc.financ18 = 2, 'Si',IF(bc.financ18 = 1, 'No', 'S/CAP') ) AS Financiamiento, " & chr(13) & chr(10)

				SQL = SQL & "COUNT( IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA')  AND bc.asie11 = '1109000000000000',bc.refe11,NULL) ) AS 'Cantidad de CG', " & chr(13) & chr(10)
				SQL = SQL & "DATE_FORMAT(MAX(bc.fech11),'%d/%m/%Y') AS 'Fecha Ultimo Movimiento', " & chr(13) & chr(10)

				SQL = SQL & "IF( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) + " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) <0,0, " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) + " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
				SQL = SQL & ")  AS 'Pagos Hechos del Saldo', " & chr(13) & chr(10)

				SQL = SQL & "IF(SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
				SQL = SQL & "<0, " & chr(13) & chr(10)
				SQL = SQL & "( ( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000100000000',bc.Saldo ,0))   / SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('4100000100000000', '4100000200080000','2111000000000000','2106000000000000'),bc.Saldo,0))  )  * (  SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) )  )  + SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000100000000',bc.Saldo ,0))  , " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000100000000',bc.Saldo ,0)) ) " & chr(13) & chr(10)
				SQL = SQL & "AS 'Honorarios del Saldo', " & chr(13) & chr(10)

				SQL = SQL & "IF(SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
				SQL = SQL & "<0, " & chr(13) & chr(10)
				SQL = SQL & "( ( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000200080000',bc.Saldo ,0))   / SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('4100000100000000', '4100000200080000','2111000000000000','2106000000000000'),bc.Saldo,0))  )  * (  SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) )  )  + SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000200080000',bc.Saldo ,0))  , " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '4100000200080000',bc.Saldo ,0)) ) " & chr(13) & chr(10)
				SQL = SQL & "AS 'Servicios Complementarios del Saldo', " & chr(13) & chr(10)

				SQL = SQL & "IF(SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|FA2' , bc.MONT11*-1,0) ) " & chr(13) & chr(10)
				SQL = SQL & "<0, " & chr(13) & chr(10)
				SQL = SQL & "( ( SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('2111000000000000','2106000000000000'),bc.Saldo ,0))   / SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('4100000100000000', '4100000200080000','2111000000000000','2106000000000000'),bc.Saldo,0))  )  * (  SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 = '1109000000000000',bc.Saldo ,0)) +SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') " & chr(13) & chr(10)

				SQL = SQL & "AND bc.asie11 IN ('1103000100000000') " & chr(13) & chr(10)

				SQL = SQL & ", bc.Saldo,0) ) )  )  + SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('2111000000000000','2106000000000000'),bc.Saldo ,0))  , " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') AND bc.asie11 IN ('2111000000000000','2106000000000000'),bc.Saldo ,0)) ) " & chr(13) & chr(10)
				SQL = SQL & "AS 'IVA Trasladado del Saldo', " & chr(13) & chr(10)

				SQL = SQL & "SUM(IF(bc.TipoCartera = 'CARTERA VENCIDA',bc.Saldo,0)) AS 'Saldo Vencido ECG', " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.TipoCartera = 'CARTERA NORMAL',bc.Saldo,0)) AS 'Saldo Vigente ECG', " & chr(13) & chr(10)

				SQL = SQL & "SUM(IF(bc.TipoCarteraRCG = 'CARTERA VENCIDA',bc.Saldo,0)) AS 'Saldo Vencido RCG', " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.TipoCarteraRCG = 'CARTERA NORMAL',bc.Saldo,0)) AS 'Saldo Vigente RCG', " & chr(13) & chr(10)

				SQL = SQL & "SUM(IF(bc.TipoCartera = 'LIQUIDADO',bc.Saldo,0)) AS 'Saldo a Favor', " & chr(13) & chr(10)
				SQL = SQL & "SUM(IF(bc.TipoCartera = 'N/A',bc.Saldo,0)) AS 'Saldo NA', " & chr(13) & chr(10)
				SQL = SQL & "SUM(bc.Saldo) AS SaldoTotal, " & chr(13) & chr(10)

				SQL = SQL & "IFNULL((SELECT SUM(   IF(bc1.conc11 IN ('ANT','CF2'),bc1.mont11, IF(bc1.conc11 IN ('FA2','CAN'),bc1.mont11*-1,0) ) ) " & chr(13) & chr(10)
				SQL = SQL & "FROM trackingbahia.Bit_Cartera_" & strOficina & " AS bc1 WHERE bc1.conc11 IN ('ANT','FA2','CF2','CAN')  AND bc1.cont11 = 'A'  AND bc1.cvecli18=bc.cvecli18  AND bc1.refe01 = bc.refe01 GROUP BY bc1.cvecli18 ),0) AS 'Anticipos Sin Aplicar', " & chr(13) & chr(10)

				SQL = SQL & "IFNULL((SELECT SUM( phe.mont21) " & chr(13) & chr(10)
				SQL = SQL & "FROM " & strOficina & "_extranet.d21paghe AS phe " & chr(13) & chr(10)
				SQL = SQL & "INNER JOIN " & strOficina & "_extranet.e21paghe AS e21 ON e21.foli21 = phe.foli21 AND YEAR(e21.fech21) = YEAR(phe.fech21) AND e21.tmov21 = phe.tmov21 " & chr(13) & chr(10)
				SQL = SQL & "INNER JOIN " & strOficina & "_extranet.c01refer AS r21 ON r21.refe01 = phe.refe21 " & chr(13) & chr(10)
				SQL = SQL & "WHERE phe.esta21 <> 'S' AND " & chr(13) & chr(10)
				SQL = SQL & "e21.esta21 NOT IN ('S','A') AND " & chr(13) & chr(10)
				SQL = SQL & "r21.clie01 = bc.cvecli18 AND phe.cgas21 =  '' AND " & chr(13) & chr(10)
				SQL = SQL & "r21.refe01 NOT IN (SELECT DISTINCT d31.refe31 FROM " & strOficina & "_extranet.d31refer AS d31 INNER JOIN " & strOficina & "_extranet.c01refer AS rr ON rr.refe01 = d31.refe31 WHERE rr.clie01 = r21.clie01) " & chr(13) & chr(10)
				SQL = SQL & "GROUP BY r21.clie01),0) AS 'Pagos Hechos Sin Aplicar', " & chr(13) & chr(10)

				SQL = SQL & "IFNULL(( ROUND(IFNULL((    (SUM(IF( (bc.facofna = '" & facofna & "'  OR bc.facofna = '')  AND bc.ASIE11 = '4100000100000000', IF(bc.CONC11 REGEXP 'FA1|SCA|DEV|CAR|FA2' , bc.MONT11, IF(bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|CF2' , bc.MONT11*-1,0)), 0)) " & chr(13) & chr(10)
				SQL = SQL & "+ SUM(IF( (bc.facofna = '" & facofna & "'  OR bc.facofna = '')  AND bc.ASIE11 = '4100000200080000',IF(bc.CONC11 REGEXP 'FA1|SCA|DEV|CAR|FA2' , bc.MONT11, IF(bc.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|CF2' , bc.MONT11*-1, 0)),0))  ) " & chr(13) & chr(10)
				SQL = SQL & "),0),2) /  IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fech11),LPAD(MONTH(rt.fech11),2,'0')  )) )  FROM trackingbahia.Bit_Cartera_" & strOficina & " AS rt WHERE rt.cvecli18 = bc.cvecli18 GROUP BY rt.cvecli18),0) ),0) AS 'Ingresos Promedio Por Mes', " & chr(13) & chr(10)

				SQL = SQL & "IFNULL(( ( IFNULL(( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdagi01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND  ri.firmae01 <> '' AND ri.cveped01 <> 'R1'),0) + " & chr(13) & chr(10)
				SQL = SQL & "IFNULL(( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdage01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND  ri.firmae01 <> '' AND ri.cveped01 <> 'R1'),0) ) / " & chr(13) & chr(10)
				SQL = SQL & "IF( IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) > " & chr(13) & chr(10)
				SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
				SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
				SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) " & chr(13) & chr(10)
				SQL = SQL & ") " & chr(13) & chr(10)
				SQL = SQL & "),0) " & chr(13) & chr(10)
				SQL = SQL & "AS 'Operaciones Promedio Por Mes', " & chr(13) & chr(10)

				SQL = SQL & "( IF( IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) > " & chr(13) & chr(10)
				SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
				SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdagi01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0), " & chr(13) & chr(10)
				SQL = SQL & "IFNULL((SELECT COUNT(DISTINCT (CONCAT(YEAR(rt.fecpag01),LPAD(MONTH(rt.fecpag01),2,'0')  )) )  FROM " & strOficina & "_extranet.ssdage01 AS rt WHERE rt.cvecli01 = bc.cvecli18 GROUP BY rt.cvecli01),0) " & chr(13) & chr(10)
				SQL = SQL & "))  AS 'Cantidad de Meses', " & chr(13) & chr(10)

				SQL = SQL & "( ( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdagi01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND ri.firmae01 <> '' AND ri.cveped01 <> 'R1') + " & chr(13) & chr(10)
				SQL = SQL & "( SELECT COUNT(DISTINCT ri.refcia01) FROM " & strOficina & "_extranet.ssdage01 AS ri WHERE ri.cvecli01 = bc.cvecli18 AND ri.firmae01 <> '' AND ri.cveped01 <> 'R1') )  AS 'Cantidad de Operaciones por Mes' " & chr(13) & chr(10)

				SQL = SQL & "FROM trackingbahia.Bit_Cartera_" & strOficina & "  AS bc " & chr(13) & chr(10)
				SQL = SQL & "WHERE " & chr(13) & chr(10)
				SQL = SQL & "bc.cont11 <> 'C' AND " & chr(13) & chr(10)
				SQL = SQL & "bc.fech11 <= '" & DateF & "' " & chr(13) & chr(10)
				SQL = SQL & condicion & chr(13) & chr(10)
				SQL = SQL & "GROUP BY bc.ccli11 " & chr(13) & chr(10)
				SQL = SQL & "HAVING SaldoTotal NOT BETWEEN -1 AND 0.1 " & chr(13) & chr(10)
		end if		
		
	case 1 'REPORTE DETALLADO
		
		if ofi = "a" then
			
			For i = 0 to 5
					
				Select Case i
						Case 0
							strOficina = "rku"
						Case 1
							strOficina = "dai"
						Case 2
							strOficina = "tol"
						Case 3
							strOficina = "sap"
						Case 4
							strOficina = "lzr"
						Case 5
							strOficina = "ceg"
				End Select
				
				'SQL = SQL & "select * " & chr(13) & chr(10)
			    'SQL = SQL & "from trackingbahia.bit_cartera_" & strOficina & " as bc " & chr(13) & chr(10)
				'SQL = SQL & "where bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') " & chr(13) & chr(10)
				'SQL = SQL & "and bc.fech11 <= '" & DateF & "' " & chr(13) & chr(10)
				'SQL = SQL & condicion & chr(13) & chr(10)
				
				SQL = SQL & " SELECT  bc.* " & chr(13) & chr(10)
				SQL = SQL & " FROM trackingbahia.bit_cartera_" & strOficina & " as bc where bc.fech11 <='" & DateF & "' " & chr(13) & chr(10)
				SQL = SQL & condicion & " and ( bc.refe11 in ( " & chr(13) & chr(10)
				SQL = SQL & " select distinct aux2.refe11 from ( " & chr(13) & chr(10)
				SQL = SQL & " select sum(aux.saldo) as tot,aux.refe11,group_concat(distinct aux.statusCG) as sta " & chr(13) & chr(10)
				SQL = SQL & " from trackingbahia.bit_cartera_" & strOficina & " as aux where aux.fech11 <=  '" & DateF & "' " & condicion &" group by  aux.refe11  " & chr(13) & chr(10)
				SQL = SQL & " having sta in ('CERRADA','CANCELADA') and (tot >=0.1 or tot <= -0.1) ) as aux2 ) " & chr(13) & chr(10)
				SQL = SQL & " or bc.statusCG in ('CON SALDO A FAVOR','OTROS CASOS','CON SALDO PENDIENTE','PENDIENTE',null) ) " & chr(13) & chr(10)
				
				
				if (i<>5) then
					SQL = SQL & "UNION ALL " & chr(13) & chr(10)
				else
					SQL = SQL & "ORDER BY 16,3 ASC "
				end if
				
			Next
			
		else
			
			Select Case ofi
					Case "r"
						strOficina = "rku"
					Case "d"
						strOficina = "dai"
					Case "t"
						strOficina = "tol"
					Case "s"
						strOficina = "sap"
					Case "l"
						strOficina = "lzr"
					Case "c"
						strOficina = "ceg"
			End Select
			
			' SQL = SQL & "select * " & chr(13) & chr(10)
			' SQL = SQL & "from trackingbahia.bit_cartera_" & strOficina & " as bc " & chr(13) & chr(10)
			' SQL = SQL & "where bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') " & chr(13) & chr(10)
			' SQL = SQL & "and bc.fech11 <= '" & DateF & "' " & chr(13) & chr(10)
			' SQL = SQL & condicion & chr(13) & chr(10) 
			' SQL = SQL & "ORDER BY 3 ASC "
			
				SQL = SQL & " SELECT  bc.* " & chr(13) & chr(10)
				SQL = SQL & " FROM trackingbahia.bit_cartera_" & strOficina & " as bc where bc.fech11 <='" & DateF & "' " & chr(13) & chr(10)
				SQL = SQL & condicion & " and ( bc.refe11 in ( " & chr(13) & chr(10)
				SQL = SQL & " select distinct aux2.refe11 from ( " & chr(13) & chr(10)
				SQL = SQL & " select sum(aux.saldo) as tot,aux.refe11,group_concat(distinct aux.statusCG) as sta " & chr(13) & chr(10)
				SQL = SQL & " from trackingbahia.bit_cartera_" & strOficina & " as aux where aux.fech11 <=  '" & DateF & "' " & condicion &" group by  aux.refe11  " & chr(13) & chr(10)
				SQL = SQL & " having sta in ('CERRADA','CANCELADA') and (tot >=0.1 or tot <= -0.1) ) as aux2 ) " & chr(13) & chr(10)
				SQL = SQL & " or bc.statusCG in ('CON SALDO A FAVOR','OTROS CASOS','CON SALDO PENDIENTE','PENDIENTE',null) ) " & chr(13) & chr(10)
			
		end if
	
	case 3 'REPORTE DE ESTADO DE CUENTA
		 'inicio de query
		 
		if ofi = "a" then
			
			For i = 0 to 5
				
				Select Case i
					Case 0
						strOficina = "rku"
						adu = "VERACRUZ"
						facofna = "0001"
					Case 1
						strOficina = "dai"
						adu = "MEXICO"
						facofna = "0005"
					Case 2
						strOficina = "tol"
						adu = "TOLUCA"
						facofna = "0010"
					Case 3
						strOficina = "sap"
						adu = "MANZANILLO"
						facofna = "0004"
					Case 4
						strOficina = "lzr"
						adu = "LAZARO CARDENAS"
						facofna = "0009"
					Case 5
						strOficina = "ceg"
						adu = "ALTAMIRA"
						facofna = "0003"
				End Select
			 
			 
				SQL = SQL & " SELECT '" & adu & "' as adu,bc.cvecli18,replace(bc.nomcli18,',','') as nomcli18,bc.plcred18,bc.fech11, bc.frec31,bc.refe11, " & chr(13) & chr(10)
				SQL = SQL & " SUM(IF(bc.conc11 in ('FA2','CF2'),bc.Saldo,0)) AS 'Anticipos',  " & chr(13) & chr(10)
				SQL = SQL & " SUM(IF(bc.TipoCarteraRCG = 'CARTERA VENCIDA',bc.Saldo,0)) AS 'Saldo Vencido RCG',  " & chr(13) & chr(10)
				SQL = SQL & " SUM(IF(bc.TipoCarteraRCG = 'CARTERA NORMAL',bc.Saldo,0)) AS 'Saldo Vigente RCG', " & chr(13) & chr(10)
				SQL = SQL & " SUM(IF(bc.TipoCartera = 'LIQUIDADO',bc.Saldo,0)) AS 'Saldo a Favor',  " & chr(13) & chr(10)
				SQL = SQL & " SUM(IF(bc.TipoCartera = 'N/A',bc.Saldo,0)) AS 'Saldo NA',  " & chr(13) & chr(10)
				SQL = SQL & " SUM(bc.Saldo) AS SaldoTotal " & chr(13) & chr(10)
				SQL = SQL & " FROM trackingbahia.bit_cartera_" & strOficina & " as bc " & chr(13) & chr(10)
				SQL = SQL & " where " & chr(13) & chr(10)
				SQL = SQL & " bc.fech11 <= '" & DateF & "'  " & condicion & chr(13) & chr(10)
				SQL = SQL & " group by bc.refe11 " & chr(13) & chr(10)
				SQL = SQL & " HAVING SaldoTotal NOT BETWEEN -1 AND 0.1" & chr(13) & chr(10)
			
				if (i<>5) then
					SQL = SQL & "UNION ALL " & chr(13) & chr(10)
				else
					SQL = SQL & "ORDER BY 3,1,2,4 "
				end if
				
			Next
		
		else
		 
			Select Case ofi
				Case "r"
					strOficina = "rku"
					adu = "VERACRUZ"
					facofna = "0001"
				Case "d"
					strOficina = "dai"
					adu = "MEXICO"
					facofna = "0005"
				Case "t"
					strOficina = "tol"
					adu = "TOLUCA"
					facofna = "0010"
				Case "s"
					strOficina = "sap"
					adu = "MANZANILLO"
					facofna = "0004"
				Case "l"
					strOficina = "lzr"
					adu = "LAZARO CARDENAS"
					facofna = "0009"
				Case "c"
					strOficina = "ceg"
					adu = "ALTAMIRA"
					facofna = "0003"
			End Select
			 
			 
			SQL = " SELECT '" & adu & "' as adu,bc.cvecli18,replace(bc.nomcli18,',','') as nomcli18,bc.plcred18,bc.fech11, bc.frec31,bc.refe11, " & chr(13) & chr(10)
			SQL = SQL & " SUM(IF(bc.conc11 in ('FA2','CF2'),bc.Saldo,0)) AS 'Anticipos',  " & chr(13) & chr(10)
			SQL = SQL & " SUM(IF(bc.TipoCarteraRCG = 'CARTERA VENCIDA',bc.Saldo,0)) AS 'Saldo Vencido RCG',  " & chr(13) & chr(10)
			SQL = SQL & " SUM(IF(bc.TipoCarteraRCG = 'CARTERA NORMAL',bc.Saldo,0)) AS 'Saldo Vigente RCG', " & chr(13) & chr(10)
			SQL = SQL & " SUM(IF(bc.TipoCartera = 'LIQUIDADO',bc.Saldo,0)) AS 'Saldo a Favor',  " & chr(13) & chr(10)
			SQL = SQL & " SUM(IF(bc.TipoCartera = 'N/A',bc.Saldo,0)) AS 'Saldo NA',  " & chr(13) & chr(10)
			SQL = SQL & " SUM(bc.Saldo) AS SaldoTotal " & chr(13) & chr(10)
			SQL = SQL & " FROM trackingbahia.bit_cartera_" & strOficina & " as bc " & chr(13) & chr(10)
			SQL = SQL & " where " & chr(13) & chr(10)
			SQL = SQL & " bc.fech11 <= '" & DateF & "'  " & condicion & chr(13) & chr(10)
			SQL = SQL & " group by bc.refe11 " & chr(13) & chr(10)
			SQL = SQL & " HAVING SaldoTotal NOT BETWEEN -1 AND 0.1" & chr(13) & chr(10)
		'fin de query
		end if
	
	end select
	
	   ' Response.Write(SQL)
	   ' Response.End
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
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


Tiporepo = Request.Form("TipRep")

if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	blnAplicaFiltro = true
end if
if blnAplicaFiltro then
	permi = " AND cvecli18 =" & strFiltroCliente
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
	' response.write(Vrfc & " | ")
	' response.write(Vckcve & " | ")
	' response.write(Vclave & " | ")
	' response.end()

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF


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
	
		Response.Addheader "Content-Disposition", "attachment;"
		Response.ContentType = "application/vnd.ms-excel"
		
		if Tiporepo = 2 Then
		
			nocolumns = 25
			info = 	"<table  width = ""2929""  border = ""1"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left"">" &_
												"REPORTE DE CARTERA GENERAL" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
												"Fecha de Corte Al " & ff &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
			header = 	"<tr class = ""boton"">" & _
		 
						celdahead("Clave del Cliente") & _
						celdahead("RFC del Cliente") & _
						celdahead("Nombre del Cliente") & _
						celdahead("Oficina") & _
						celdahead("Dias de Credito") & _
						celdahead("Financiamiento") & _
						celdahead("Cantidad de CG") & _
						celdahead("Fecha Ultimo Movimiento") & _
						celdahead("Pagos Hechos del Saldo") & _
						celdahead("Honorarios del Saldo") & _
						celdahead("Servicios Complementarios del Saldo") & _
						celdahead("IVA Trasladado del Saldo") & _
						celdahead("Saldo Vencido ECG") & _
						celdahead("Saldo Vigente ECG") & _
						celdahead("Saldo Vencido RCG") & _
						celdahead("Saldo Vigente RCG") & _
						celdahead("Saldo a Favor") & _
						celdahead("Saldo NA") & _
						celdahead("Saldo Total") & _
						celdahead("Anticipos Sin Aplicar") & _
						celdahead("Pagos Hechos Sin Aplicar") & _
						celdahead("Ingresos Promedio Por Mes") & _
						celdahead("Operaciones Promedio Por Mes") & _
						celdahead("Cantidad de Meses") & _
						celdahead("Cantidad de Operaciones por Mes")

			header = header &	"</tr>"
		
			'celdahead("Proveedor") &_
			''celdahead("Descripcion de la Mercancia") &_
									
			
			Do Until RSops.EOF
 
				datos = datos &	"<tr>" &_
				
					celdadatos(RSops.Fields.Item("Clave del Cliente").Value) & _
					celdadatos(RSops.Fields.Item("RFC del Cliente").Value) & _
					celdadatos(RSops.Fields.Item("Nombre del Cliente").Value) & _
					celdadatos(RSops.Fields.Item("Oficina").Value) & _
					celdadatos(RSops.Fields.Item("Dias de Credito").Value) & _
					celdadatos(RSops.Fields.Item("Financiamiento").Value) & _
					celdadatos(RSops.Fields.Item("Cantidad de CG").Value) & _
					celdadatos(RSops.Fields.Item("Fecha Ultimo Movimiento").Value) & _
					celdadatos(RSops.Fields.Item("Pagos Hechos del Saldo").Value) & _
					celdadatos(RSops.Fields.Item("Honorarios del Saldo").Value) & _
					celdadatos(RSops.Fields.Item("Servicios Complementarios del Saldo").Value) & _
					celdadatos(RSops.Fields.Item("IVA Trasladado del Saldo").Value) & _
					celdadatos(RSops.Fields.Item("Saldo Vencido ECG").Value) & _
					celdadatos(RSops.Fields.Item("Saldo Vigente ECG").Value) & _
					celdadatos(RSops.Fields.Item("Saldo Vencido RCG").Value) & _
					celdadatos(RSops.Fields.Item("Saldo Vigente RCG").Value) & _
					celdadatos(RSops.Fields.Item("Saldo a Favor").Value) & _
					celdadatos(RSops.Fields.Item("Saldo NA").Value) & _
					celdadatos(RSops.Fields.Item("SaldoTotal").Value) & _
					celdadatos(RSops.Fields.Item("Anticipos Sin Aplicar").Value) & _
					celdadatos(RSops.Fields.Item("Pagos Hechos Sin Aplicar").Value) & _
					celdadatos(RSops.Fields.Item("Ingresos Promedio Por Mes").Value) & _
					celdadatos(RSops.Fields.Item("Operaciones Promedio Por Mes").Value) & _
					celdadatos(RSops.Fields.Item("Cantidad de Meses").Value) & _
					celdadatos(RSops.Fields.Item("Cantidad de Operaciones por Mes").Value)
				
				datos = datos &	"</tr>"
								
				Rsops.MoveNext()
			Loop
			
		else
		
			nocolumns = 20
			info = 	"<table  width = ""2929""  border = ""1"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left"">" &_
												"REPORTE DE CARTERA DETALLADO" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
												"Fecha de Corte Al " & ff &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
			header = 	"<tr class = ""boton"">" & _
		 
						celdahead("Fecha del Movimiento") & _
						celdahead("Asiento") & _
						celdahead("Cuenta de Gastos") & _
						celdahead("Concepto") & _
						celdahead("Usuario Captura") & _
						celdahead("Folio") & _
						celdahead("Oficina") & _
						celdahead("Referencia") & _
						celdahead("RFC") & _
						celdahead("Cliente") & _
						celdahead("Clave de Cliente") & _
						celdahead("Plazo de Credito") & _
						celdahead("Cuenta Contable") & _
						celdahead("Tipo de Financiamiento") & _
						celdahead("Fecha de Recepcion de CG") & _
						celdahead("Saldo") & _
						celdahead("Status de Cuenta de Gastos") & _
						celdahead("Dias Pendientes") & _
						celdahead("Tipo de Cartera (Recepcion de la Cuenta de Gastos)") & _
						celdahead("Tipo de Cartera (Emision de la Cuenta de Gastos)")


			header = header &	"</tr>"
		
			'celdahead("Proveedor") &_
			''celdahead("Descripcion de la Mercancia") &_
									

			Do Until RSops.EOF
 
				datos = datos &	"<tr>" &_
				
					celdadatos(RSops.Fields.Item("fech11").Value) & _
					celdadatos(cstr(RSops.Fields.Item("Asie11").Value)) & _
					celdadatos(RSops.Fields.Item("refe11").Value) & _
					celdadatos(RSops.Fields.Item("conc11").Value) & _
					celdadatos(RSops.Fields.Item("user11").Value) & _
					celdadatos(RSops.Fields.Item("d11_folres11").Value) & _
					celdadatos(RSops.Fields.Item("facofna").Value) & _
					celdadatos(RSops.Fields.Item("refe01").Value) & _
					celdadatos(RSops.Fields.Item("cvecli18").Value) & _
					celdadatos(RSops.Fields.Item("rfccli18").Value) & _
					celdadatos(RSops.Fields.Item("nomcli18").Value) & _
					celdadatos(RSops.Fields.Item("plcred18").Value) & _
					celdadatos(RSops.Fields.Item("CuentaContable").Value) & _
					celdadatos(RSops.Fields.Item("financ18").Value) & _
					celdadatos(RSops.Fields.Item("frec31").Value) & _
					celdadatos(RSops.Fields.Item("Saldo").Value) & _
					celdadatos(RSops.Fields.Item("statusCG").Value) & _
					celdadatos(RSops.Fields.Item("DiasPendientes").Value) & _
					celdadatos(RSops.Fields.Item("TipoCarteraRCG").Value) & _
					celdadatos(RSops.Fields.Item("TipoCartera").Value)

				
				datos = datos &	"</tr>"
								
				Rsops.MoveNext()
			Loop
			
		End If

	' Response.Write(info & header & datos & "</table><br>" & prom)
	' Response.End()
	html = info & header & datos & "</table><br>"
	
	
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
	if Vckcve = 0 then
		if Vrfc <> "0" then
			condicion = "AND rfccli18 = '" & Vrfc & "' "
		else
			'condicion = "AND i.rfccli01 = 'UME651115N48' "
			condicion = " "
		end if
	else
		if Vclave <> "Todos" Then
			condicion = "AND cvecli18 = " & Vclave & " "
		Else
			permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli18")
			condicion = permi
			condicion = "AND " & condicion
			if condicion = "AND cvecli18=0 " Then
				condicion = ""
			end if
		End If
	end if
	filtro = condicion
end function

function GeneraSQL
	SQL = ""
	condicion = filtro
	
	if Tiporepo = 2 Then
	
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
					
					SQL = SQL & "SELECT bc.cvecli18 as 'Clave del Cliente',bc.rfccli18 'RFC del Cliente',bc.nomcli18 as 'Nombre del Cliente','" & adu & "' AS Oficina, " & chr(13) & chr(10)
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
					SQL = SQL & "GROUP BY bc.cvecli18 " & chr(13) & chr(10)
					SQL = SQL & "HAVING SaldoTotal NOT BETWEEN -1 AND 0.1 " & chr(13) & chr(10)
					
					if (i<>5) then
						SQL = SQL & "UNION ALL " & chr(13) & chr(10)
					else
						SQL = SQL & "ORDER BY 4 ASC "
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

				SQL = SQL & "SELECT bc.cvecli18 as 'Clave del Cliente',bc.rfccli18 'RFC del Cliente',bc.nomcli18 as 'Nombre del Cliente','" & adu & "' AS Oficina, " & chr(13) & chr(10)
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
				SQL = SQL & "GROUP BY bc.cvecli18 " & chr(13) & chr(10)
				SQL = SQL & "HAVING SaldoTotal NOT BETWEEN -1 AND 0.1 " & chr(13) & chr(10)
				
		end if
		
	else
		
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
				
				SQL = SQL & "select * " & chr(13) & chr(10)
				SQL = SQL & "from trackingbahia.bit_cartera_" & strOficina & " as bc " & chr(13) & chr(10)
				SQL = SQL & "where bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') " & chr(13) & chr(10)
				SQL = SQL & condicion & chr(13) & chr(10)
				
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
			
			SQL = SQL & "select * " & chr(13) & chr(10)
			SQL = SQL & "from trackingbahia.bit_cartera_" & strOficina & " as bc " & chr(13) & chr(10)
			SQL = SQL & "where bc.StatusCG NOT IN ('N/A','CERRADA','CANCELADA') " & chr(13) & chr(10)
			SQL = SQL & condicion & chr(13) & chr(10) 
			SQL = SQL & "ORDER BY 3 ASC "
			
		end if
		
	end if
	   ' Response.Write(SQL)
	   ' Response.End
	GeneraSQL = SQL
end function




%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE CARTERA .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>
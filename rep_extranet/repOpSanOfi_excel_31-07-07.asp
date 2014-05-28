<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->


<%
' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE


Response.Addheader "Content-Disposition", "attachment;"
Response.ContentType = "application/vnd.ms-excel"


Server.ScriptTimeOut=100000




strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
permi2 = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")




if not permi2 = "" then
  permi2 = "  and (" & permi2 & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi2 = " AND B.clie31 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi2 = ""
end if







if not permi = "" then
  permi = "  and (" & permi & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi = " AND cvecli01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if

'response.Write("Permisos="&permi)



strDateIni=""
strDateFin=""
strTipoPedimento= ""
strCodError = "0"

strDateIni=trim(request.Form("txtDateIni"))
strDateFin=trim(request.Form("txtDateFin"))
'*******************************************************
' Si es Impo o Expo
strTipoPedimento=trim(request.Form("rbnTipoDate"))
'*******************************************************

'***************************************************************************************************************
strDescripcion=trim(request.Form("txtDescripcion"))
strDateIni2=trim(request.Form("txtDateIni2"))
strDateFin2=trim(request.Form("txtDateFin2"))
strTipoPedimento2=trim(request.Form("rbnTipoDate2"))

strTipoFiltro=trim(request.Form("TipoFiltro"))

if not isdate(strDateIni) then
	strCodError = "5"
end if
if not isdate(strDateFin) then
	strCodError = "6"
end if
if strDateIni="" or strDateFin="" then
	strCodError = "1"
end if


if strCodError = "0" then

strHTML = ""
strDateIni=trim(request.Form("txtDateIni"))
strDateFin=trim(request.Form("txtDateFin"))
strTipoPedimento=trim(request.Form("rbnTipoDate"))
strUsuario = request.Form("user")
strTipoUsuario = request.Form("TipoUser")

tmpTipo = ""
strSQL = ""


if strTipoPedimento  = "1" then
   tmpTipo = "IMPORTACION"
   strSQL = "SELECT tipopr01, " & _
            "       valmer01, " & _
            "       factmo01, " & _
            "       p_dta101, " & _
            "       t_reca01, " & _
            "       i_dta101, " & _
            "       cvecli01, " & _
            "       refcia01, " & _
            "       fecpag01, " & _
            "       valfac01, " & _
            "       fletes01, " & _
            "       segros01, " & _
            "       cvepvc01, " & _
            "       tipcam01, " & _
            "       patent01, " & _
            "       numped01, " & _
            "       totbul01, " & _
            "       cveped01, " & _
            "       cveadu01, " & _
            "       desf0101, " & _
            "       nompro01, " & _
            "       cvepod01, " & _
            "       nombar01, " & _
            "       tipopr01, " & _
            "       fecent01, " & _
            "       if(alea01=1,'ROJO',if(alea01=2, 'VERDE',' ' ) ) as semaforo, " & _
            "       CVEMTS01 as destino " & _
            "FROM ssdagi01 ,c01refer " & _
            "WHERE refcia01=refe01 and " & _
            "      fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
            "      fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & _
                   Permi & " and " & _
            "      firmae01 !='' " & _
            "order by refcia01"

'            "       formfa01 as incoterm  " & _

            StrSQLPH=" SELECT distinct conc21,desc21 " & _
                     " FROM SSDAGI01,                "  & _
                     "      d21paghe,                " & _
                     "      e21paghe,                " & _
                     "      c21paghe                 " & _
                     " WHERE refcia01                  =  d21paghe.refe21 and " & _
                     "       YEAR(d21paghe.fech21)     =  YEAR(e21paghe.fech21) and " & _
                     "       d21paghe.foli21           =  e21paghe.foli21 and " & _
                     "       d21paghe.tmov21           =  e21paghe.tmov21 and " & _
                     "       e21paghe.fech21           =  d21paghe.fech21 AND " & _
                     "       e21paghe.esta21          !=  'S' and " & _
                     "       conc21                    =  clav21 and              " & _
                     "       fecpag01                 >=  '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     "       fecpag01                 <='"&FormatoFechaInv(strDateFin)&"' " & _
                             Permi & " and " & _
                     "       firmae01 !='' " & _
                     " ORDER BY CONC21  "



end if
if strTipoPedimento  = "2" then
   tmpTipo = "EXPORTACION"
   strSQL = "SELECT tipopr01, " & _
            "       ' ' as  valmer01, " & _
            "       factmo01, " & _
            "       p_dta101, " & _
            "       t_reca01, " & _
            "       i_dta101, " & _
            "       cvecli01, " & _
            "       refcia01, " & _
            "       fecpag01, " & _
            "       valfac01, " & _
            "       fletes01, " & _
            "       segros01, " & _
            "       cvepvc01, " & _
            "       tipcam01, " & _
            "       patent01, " & _
            "       numped01, " & _
            "       totbul01, " & _
            "       cveped01, " & _
            "       cveadu01, " & _
            "       desf0101, " & _
            "       nompro01, " & _
            "       cvepod01, " & _
            "       nombar01, " & _
            "       tipopr01, " & _
            "       fecpre01, " & _
            "       if(alea01=1,'ROJO',if(alea01=2, 'VERDE',' ' ) ) as semaforo, " & _
            "       CVEMTS01 as destino " & _
            "FROM ssdage01 ,c01refer " & _
            "WHERE refcia01=refe01 and " & _
            "      fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
            "      fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & _
                   Permi & " and " & _
            " firmae01 !='' " & _
            "order by refcia01"

            StrSQLPH=" SELECT distinct conc21,desc21 " & _
                     " FROM SSDAGE01,                "  & _
                     "      d21paghe,                " & _
                     "      e21paghe,                " & _
                     "      c21paghe                 " & _
                     " WHERE refcia01                  =  d21paghe.refe21 and " & _
                     "       YEAR(d21paghe.fech21)     =  YEAR(e21paghe.fech21) and " & _
                     "       d21paghe.foli21           =  e21paghe.foli21 and " & _
                     "       d21paghe.tmov21           =  e21paghe.tmov21 and " & _
                     "       e21paghe.fech21           =  d21paghe.fech21 AND " & _
                     "       e21paghe.esta21          !=  'S' and " & _
                     "       conc21                    =  clav21 and              " & _
                     "       fecpag01                 >=  '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     "       fecpag01                 <='"&FormatoFechaInv(strDateFin)&"' " & _
                             Permi & " and " & _
                     "       firmae01 !='' " & _
                     " ORDER BY CONC21  "

end if

'response.Write("Query="&strSQL)


if not trim(strSQL)="" then
		Set RsRep = Server.CreateObject("ADODB.Recordset")
		RsRep.ActiveConnection = MM_EXTRANET_STRING
		RsRep.Source = strSQL
		RsRep.CursorType = 0
		RsRep.CursorLocation = 2
		RsRep.LockType = 1
		RsRep.Open()

	if not RsRep.eof then
  ' Comienza el HTML, se pintan los titulos de las columnas
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE OPERACIONES DE SANOFI DE " & tmpTipo & " </p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>GRUPO REYES KURI, S.C. </p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>" & strDateIni & " al " & strDateFin & "</p></font></strong>"
     strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	   strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)


		'strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de Pedimento</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & Hd1 & "</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No.Pitex</font></td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</font></td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observaciones</font></td>" & chr(13) & chr(10)
    'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</font></td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No. de Contenedores</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Bultos</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave de Documento</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Origen/Destino</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Guia/B.L.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Buque</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Entrada/Presentacion</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Pago</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Diferencia</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de cambio</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Vendedor/Comprador</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val. Dol. Fact.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val.Mcia. M.N.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fletes</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Seguros</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Factura</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cve. Unidad Fact.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Tarifa</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Unidad Tarifa</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Dol.(Fraccion)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Mcia.(Fraccion)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana(Fraccion)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago DTA(1)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe DTA(1)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa Advalorem IGI</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago Advalorem IGI(1)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Advalorem IGI(1)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA(1)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa de Recargos(1)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Recargos</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuotas Compensatorias</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Efectivo</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Otros</td>" & chr(13) & chr(10)
    'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Identificadores</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de la C.G.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pagos Hechos</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Servicios Complemetarios</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Anticipos</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Saldo de la C.G.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contacto</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C. del Cliente</td>" & chr(13) & chr(10)
    'strHTML = strHTML & "</tr>"& chr(13) & chr(10)

		strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de Pedimento</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Agente Aduanal</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedido</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Destino</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observaciones</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Codigo Parte</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion Factura</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Incoterm</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No. de Contenedores</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Bultos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave de Documento</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Origen/Destino</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Guia/B.L.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Buque</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Entrada/Presentacion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Pago</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Semaforo</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de cambio</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Vendedor/Comprador</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val. Dol. Fact.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val.Mcia. M.N.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fletes</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Seguros</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Factura</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cve. Unidad Fact.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Tarifa</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Unidad Tarifa</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Dol.(Fraccion)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Mcia.(Fraccion)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana(Fraccion)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago DTA(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe DTA(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa Advalorem IGI</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago Advalorem IGI(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Advalorem IGI(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Efectivo</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Otros</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Identificadores</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Permisos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pagos Hechos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Servicios Complemetarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Anticipos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Saldo de la C.G.</td>" & chr(13) & chr(10)

    ' Desglozar los pagos hechos
    	Set RsRepPHConceptos = Server.CreateObject("ADODB.Recordset")
		  RsRepPHConceptos.ActiveConnection = MM_EXTRANET_STRING
		  RsRepPHConceptos.Source = StrSQLPH
		  RsRepPHConceptos.CursorType = 0
		  RsRepPHConceptos.CursorLocation = 2
		  RsRepPHConceptos.LockType = 1
		  RsRepPHConceptos.Open()
      IntLonPH=0




      While NOT RsRepPHConceptos.EOF
         IntLonPH=IntLonPH + 1
         RsRepPHConceptos.movenext
      wend
      'RsRepPHConceptos.close

      Dim ConceptosPagosHechos()
      redim ConceptosPagosHechos(IntLonPH)

      IntIndice=0
      'RsRepPHConceptos.Open()
      RsRepPHConceptos.MoveFirst
      While NOT RsRepPHConceptos.EOF
         strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & RsRepPHConceptos.Fields.Item("desc21").Value & "</td>" & chr(13) & chr(10)
         ConceptosPagosHechos(IntIndice) = RsRepPHConceptos.Fields.Item("conc21").Value
         RsRepPHConceptos.movenext
         IntIndice=IntIndice + 1
      wend

      RsRepPHConceptos.close

      Set RsRepPHConceptos = Nothing

      'response.Write(ConceptosPagosHechos(0) )
      'response.Write(ConceptosPagosHechos(1) )

		  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contacto</td>" & chr(13) & chr(10)
		  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C. del Cliente</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)




'*************************************************************


	While NOT RsRep.EOF

    'Se asigna el nombre de la referencia
     strRefer = RsRep.Fields.Item("refcia01").Value

     'Checa si la referencia es rectificada o rectificacion
     'Si es una de ellas lo almacena en ObservRect

     ObservRect = ""
     strRect = ""
     strRect = RegresaRect(strRefer,"Rectificado")
     if not strRect = ""  then
        ObservRect=strRect
     end if
     strRect = RegresaRect(strRefer,"Rectificacion")
     if not strRect = ""  then
        ObservRect=strRect
     end if

     dim intContReg
     dim intContCtas
     dim strCtaGas
	   dim intContCG

	   intContCtas=1
     intContReg=1
     strCtaGas=""
	   strfech31=""
	   dblsuph31=0
	   dblcoad31=0
	   dblcsce31=0
	   dblchon31=0
	   dblpiva31=0
	   dblanti31=0
	   dblsald31=0

        ' Aqui se buscan los datos de la cuenta de gastos
        strSQL1="select A.cgas31,B.fech31,B.suph31, B.coad31, B.csce31, B.chon31, B.piva31, B.anti31,B.sald31 from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"' " & permi2


'response.Write(strSQL1)
'response.end

        Set RsRep1 = Server.CreateObject("ADODB.Recordset")
        RsRep1.ActiveConnection = MM_EXTRANET_STRING
        RsRep1.Source = strSQL1
        RsRep1.CursorType = 0
        RsRep1.CursorLocation = 2
        RsRep1.LockType = 1
        RsRep1.Open()

        'Si no tiene una cuenta de gastos, igual se mandan a desplegar los datos de la referencia
       if RsRep1.eof then
          strHTML=DespliegaRepDesgRef(strRefer, strCtaGas)

       else
        'Si tiene varias cuentas de gastos se repiten los datos de la referencia y se despliegan los distintos datos de las diferentes cuentas de gastos
          while not RsRep1.eof
           intContCG=1
           strCtaGas=RsRep1.Fields.Item("cgas31").Value
           strfech31=RsRep1.Fields.Item("fech31").Value
           dblsuph31=cdbl(RsRep1.Fields.Item("suph31").Value)
           dblcoad31=cdbl(RsRep1.Fields.Item("coad31").Value)
           dblcsce31=cdbl(RsRep1.Fields.Item("csce31").Value)
           dblchon31=cdbl(RsRep1.Fields.Item("chon31").Value)
           dblpiva31=cdbl(RsRep1.Fields.Item("piva31").Value)
           dblanti31=cdbl(RsRep1.Fields.Item("anti31").Value)
           dblsald31=cdbl(RsRep1.Fields.Item("sald31").Value)
           strHTML=DespliegaRepDesgRef(strRefer, strCtaGas)
           intContCtas=intContCtas + 1
           RsRep1.movenext
          wend
       end if
		RsRep1.close
        Set RsRep1=Nothing
   RsRep.movenext
   Wend

'************************

   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   'Se pinta todo el HTML formado
   response.Write(strHTML)
   end if
   RsRep.close
   Set RsRep = Nothing




   if strHTML = "" then
      strHTML = "NO EXISTEN REGISTROS"
      response.Write(strHTML)
   end if
 else
   strHTML = "NO EXISTEN REGISTROS"
   response.Write(strHTML)
end if
%>




<%
'Funcion que va elaborando el reporte desglosado de referencias y devuelve el HTML
function DespliegaRepDesgRef(pRefer, pCtaGas)
      'Sus parametros son
      'pRefer    Referencia
      'pCtaGas  Cuenta de Gastos
       dim dblEfectivo
       dim dblOtros

        strSQL3="select fpagoi36, " & _
                "        import36 " & _
                "from sscont36    " & _
                "where refcia36='"&pRefer&"'"

			Set RsRep3 = Server.CreateObject("ADODB.Recordset")
			RsRep3.ActiveConnection = MM_EXTRANET_STRING
			RsRep3.Source = strSQL3
			RsRep3.CursorType = 0
			RsRep3.CursorLocation = 2
			RsRep3.LockType = 1
			RsRep3.Open()

      if not RsRep3.eof then
      ' Aqui se obtienen los campos de Suma de Efectivo y Suma de Otros
        dblEfectivo=0
        dblOtros=0
         while not RsRep3.eof
          If cdbl(RsRep3.Fields.Item("fpagoi36").Value)=0 then
              dblEfectivo=dblEfectivo+cdbl(RsRep3.Fields.Item("import36").Value)  'Sumamos el efectivo para una referencia
          else
              dblOtros=dblOtros+cdbl(RsRep3.Fields.Item("import36").Value)         'Sumamos los otros conceptos para una referencia
          end if
			  RsRep3.movenext
         wend
         RsRep3.close
         Set RsRep3=Nothing
      end if




      strSQLIncoterm=" SELECT refcia39, " & _
                     "      terfac39    " & _
                     " FROM ssfact39    " & _
                     " WHERE refcia39='"&pRefer&"'"

			Set RsRepIncoterm = Server.CreateObject("ADODB.Recordset")
			RsRepIncoterm.ActiveConnection = MM_EXTRANET_STRING
			RsRepIncoterm.Source = strSQLIncoterm
			RsRepIncoterm.CursorType = 0
			RsRepIncoterm.CursorLocation = 2
			RsRepIncoterm.LockType = 1
			RsRepIncoterm.Open()
      StrIncoterm=" "

      if not RsRepIncoterm.eof then
          StrIncoterm=RsRepIncoterm.Fields.Item("terfac39").Value
      end if
      RsRepIncoterm.close
      Set RsRepIncoterm=Nothing


      'Aqui se obtienen las fracciones por referencia
			strSQL2=" select ordfra02," & _
              "        fraarn02," & _
              "        d_mer102," & _
              "        cancom02," & _
              "        cancom02," & _
              "        u_medc02," & _
              "        cantar02," & _
              "        u_medt02," & _
              "        vmerme02," & _
              "        vaduan02," & _
              "        tasadv02," & _
              "        p_adv102," & _
              "        i_adv102," & _
              "        i_iva102," & _
              "        i_adv102," & _
              "        i_adv202," & _
              "        i_cc0102," & _
              "        i_cc0202 " & _
              " from ssfrac02   " & _
              " where refcia02='"&pRefer&"'"

			Set RsRep2 = Server.CreateObject("ADODB.Recordset")
			RsRep2.ActiveConnection = MM_EXTRANET_STRING
			RsRep2.Source = strSQL2
			RsRep2.CursorType = 0
			RsRep2.CursorLocation = 2
			RsRep2.LockType = 1
			RsRep2.Open()




			if not RsRep2.eof then
      	 While not RsRep2.eof


            'Recorremos las Fracciones
				    strHTML = strHTML&"<tr>" & chr(13) & chr(10)
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&TipoOper(RsRep.Fields.Item("tipopr01").Value)&"&nbsp;</font></td>" & chr(13) & chr(10) 'Tipo de Pedimento

				    cmp1 = ""
				    if strTipoUsuario = MM_Cod_Cliente_Division then 'Para clientes con division
				       cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"division18")
				    else
				       cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"nomcli18")
				    end if

            strSQL= "select REFCIA12," & _
                    "       ordfra12," & _
                    "       cveide12," & _
                    "       numper12 " & _
                    "from ssipar12   " & _
                    "where ordfra12 =" &  RsRep2.Fields.Item("ordfra02").Value & " and " & _
                    " refcia12='" &  pRefer & "'"

            Set RsIdent = Server.CreateObject("ADODB.Recordset")
			      RsIdent.ActiveConnection = MM_EXTRANET_STRING
			      RsIdent.Source = strSQL
			      RsIdent.CursorType = 0
			      RsIdent.CursorLocation = 2
			      RsIdent.LockType = 1
			      RsIdent.Open()
            strIdentificadores=""
            strpermisos= " "

			      if not RsIdent.eof then
              While not RsIdent.eof
                 strIdentificadores = strIdentificadores  & "  " & RsIdent.Fields.Item("cveide12").Value  & " &nbsp; "
                 strpermisos= strpermisos & "  " & RsIdent.Fields.Item("numper12").Value  & " &nbsp; "
                 RsIdent.movenext
              wend
            end if
            RsIdent.close
            set RsIdent = Nothing


            ' IIF(ALLTRIM(ADUANASECCION)=='430',
            ' "GRUPO REYES KURI, S.C.", IIF(ALLTRIM(ADUANASECCION)=='160',"SERVICIOS ADUANALES DEL PACIFICO S.C.", "DESPACHOS AEREOS INTEGRADOS, S.C." ) )  as AgenciaAduanal


            Strpatente=RsRep.Fields.Item("patent01").Value
            Stragenteaduanal=" "
            if(Strpatente="3857") then
                 Stragenteaduanal= "RAFAEL MENDOZA DIAZ BARRIGA"
            else
               if(Strpatente="3883") then
                   Stragenteaduanal= "MA SUSANA DE GPE FRICKE URQUIOLA        "
               else
                   if(Strpatente="3210") then
                      Stragenteaduanal= "LIC. ROLANDO REYES KURI "
                   else
                       if(Strpatente="3044") then
                          Stragenteaduanal= "LIC. CARLOS HUMBERTO ZESATI ANDRADE"
                       else
                          Stragenteaduanal= " "
                       end if
                   end if
               end if
            end if


            strSQLArt= "SELECT REFE05," & _
                       "       PEDI05," & _
                       "       CPRO05," & _
                       "       DESC05 " & _
                       "from d05artic " & _
                       "where refe05 ='" &  pRefer & "' AND " & _
                       "      FRAC05 ='" &  RsRep2.Fields.Item("fraarn02").Value &"' AND  " & _
                       "      AGRU05 ="  &  RsRep2.Fields.Item("ordfra02").Value

            Set RsArticulos = Server.CreateObject("ADODB.Recordset")
			      RsArticulos.ActiveConnection = MM_EXTRANET_STRING
			      RsArticulos.Source = strSQLArt
			      RsArticulos.CursorType = 0
			      RsArticulos.CursorLocation = 2
			      RsArticulos.LockType = 1
			      RsArticulos.Open()
            strPedido=""
            strCodigoParte= ""
            strDescFactura= ""
			      if not RsArticulos.eof then
                 strPedido= RsArticulos.Fields.Item("PEDI05").Value
                 strCodigoParte= RsArticulos.Fields.Item("CPRO05").Value
                 strDescFactura= RsArticulos.Fields.Item("DESC05").Value
            end if
            RsArticulos.close
            set RsArticulos = Nothing





				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pRefer&"</font></td>" & chr(13) & chr(10)  'Referencia
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Stragenteaduanal&"</font></td>" & chr(13) & chr(10) 'Agente Aduanal
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strPedido&"</font></td>" & chr(13) & chr(10) 'Pedido
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("destino").Value&"</font></td>" & chr(13) & chr(10) 'Destino
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&ObservRect&"</font></td>" & chr(13) & chr(10) 'Observaciones, si es rectificado o rectificacion
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strCodigoParte&"</font></td>" & chr(13) & chr(10) 'Codigo Parte
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strDescFactura&"</font></td>" & chr(13) & chr(10) 'Descripcion Factura
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&StrIncoterm&"</font></td>" & chr(13) & chr(10) 'Incoterm
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("patent01").Value&"</font></td>" & chr(13) & chr(10) 'Patente
			      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
            if intContReg=1 then
               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaNoCont(pRefer)&"</font></td>" & chr(13) & chr(10) 'Numero de Contenedores
            else
               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Numero de Contenedores
            end if
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("totbul01").Value&"</font></td>" & chr(13) & chr(10) 'Total de Bultos
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveped01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de pedimento
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveadu01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de aduana
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("desf0101").Value&"</font></td>" & chr(13) & chr(10) 'Facturas
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nompro01").Value&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep.Fields.Item("cvepod01").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Origen/Destino
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaGuia(pRefer)&"</font></td>" & chr(13) & chr(10) 'Si tiene guia
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nombar01").Value&"</font></td>" & chr(13) & chr(10) 'Buque,Nombre del barco
			      if RsRep.Fields.Item("tipopr01").Value="1" then 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
					     FecPed = RsRep.Fields.Item("fecent01").Value
			      else
			         FecPed = RsRep.Fields.Item("fecpre01").Value
			      end if
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FecPed&"</font></td>" & chr(13) & chr(10) 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha de pago
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("semaforo").Value&"</font></td>" & chr(13) & chr(10) 'Semaforo
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tipcam01").Value&"</font></td>" & chr(13) & chr(10) 'Tipo de Cambio
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep.Fields.Item("cvepvc01").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Vendedor/Comprador
            if intContReg=1 then
               if RsRep.Fields.Item("tipopr01").Value="1" then 'valor Factura en Dolares si es Impo
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
               end if
               if RsRep.Fields.Item("tipopr01").Value="2" then 'valor Factura en Dolares si es Expo
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valfac01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valfac01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
               end if
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoSumaValoresFracc("vaduan02", pRefer)&"</font></td>" & chr(13) & chr(10) 'Valor Aduana
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fletes01").Value&"</font></td>" & chr(13) & chr(10) 'fletes
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("segros01").Value&"</font></td>" & chr(13) & chr(10) 'Seguros
            else
               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val. Dol. Fact.
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val.Mcia. M.N.
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Valor Aduana
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'fletes
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Seguros
            end if
		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("fraarn02").Value&"</font></td>" & chr(13) & chr(10) 'Fraccion Arancearia
		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("d_mer102").Value&"</font></td>" & chr(13) & chr(10) 'Descripcion de la Mercancia
            if intContCtas=1 then
			         strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cancom02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
     	      else
	   	         if pCtaGas="" and intContReg=1 then
		 	             strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cancom02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
		           end if
		 	         strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) ''Cant. Factura
	          end if
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("u_medc02").Value&"</font></td>" & chr(13) & chr(10) 'Cve. Unidad Fact.
            if intContCtas=1 then
			         strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		        else
		          if pCtaGas="" and intContReg=1 then
		             strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
              end if
		          strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		        end if
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("u_medt02").Value&"</font></td>" & chr(13) & chr(10) 'Unidad Tarifa
            if intContCtas=1 then
                strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
                strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
                strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
            else
              if pCtaGas="" and intContReg=1 then
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
              else
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
              end if
		        end if
			      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("p_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago DTA(1)

	 	        if intContReg=1 then
			          strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Importe DTA(1)
		        else
			          strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Importe DTA(1)
            end if

            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("tasadv02").Value&"</font></td>" & chr(13) & chr(10) 'Tasa Advalorem IGI
		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago Advalorem IGI(1)

            if intContCtas=1 then
        	      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
		        else
			         if pCtaGas="" and intContReg=1 then
        		      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		    	        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
			         else
          		    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		  		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA(1)
        	     end if
		        end if
            if intContReg=1 then
                strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblEfectivo&"</font></td>" & chr(13) & chr(10) 'Efectivo
		            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblOtros&"</font></td>" & chr(13) & chr(10) 'Otros
            else
		            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Efectivo
                strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Otros
            end if
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strIdentificadores&"</font></td>" & chr(13) & chr(10) 'Identificadores
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strpermisos &"</font></td>" & chr(13) & chr(10) 'Permisos
            if not pCtaGas ="" then
               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pCtaGas&"</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
  		         strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strfech31&"</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.
               if intContCG=1 then
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsuph31 + dblcoad31&"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		              strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblcsce31&"</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblchon31&"</font></td>" & chr(13) & chr(10) 'Honorarios
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&(dblcsce31 + dblchon31)*(dblpiva31/100)&"</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblanti31&"</font></td>" & chr(13) & chr(10) 'Anticipos
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsald31&"</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
               else
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
               end if
               intContCG=intContCG + 1
            else
               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">/ /</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
		           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
            end if

             '--------------------------------------------
             ' Desglozar los pagos hechos
             '--------------------------------------------

              'response.Write(ConceptosPagosHechos(0))
              'response.Write(ConceptosPagosHechos(1))
              'response.Write("PagosHechos")
              'for each valor in ConceptosPagosHechos
             '    response.Write("PH"&valor)
              'next

            if intContReg=1 then
               StrQryConcppH= ""
               StrQryConcppH= " SELECT  "
               for inti=0 to (IntLonPH-1)
                 '       strHTML = strHTML & "<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & ConceptosPagosHechos(inti) & "</FONT></td>" & chr(13) & chr(10) 'Pagos hechos desglozados
                  IF inti=0 THEN
                     StrQryConcppH = StrQryConcppH & "sum(if(conc21="& ConceptosPagosHechos(inti) &",(if(trim(e21paghe.deha21)='A',d21paghe.mont21,(d21paghe.mont21)*-1)  ),0 ) ) as var"&CStr(inti)
                  ELSE
                     StrQryConcppH = StrQryConcppH & "," & "sum(if(conc21="& ConceptosPagosHechos(inti) &",(if(trim(e21paghe.deha21)='A',d21paghe.mont21,(d21paghe.mont21)*-1)  ),0 ) ) as var"&CStr(inti)
                  END IF
               next
               StrQryConcppH = StrQryConcppH & "  from  d21paghe, e21paghe " & _
                                               "  where  d21paghe.refe21  = '" &  pRefer & "' AND " & _
                                               "         YEAR(d21paghe.fech21) = YEAR(e21paghe.fech21)  and " & _
                                               "         e21paghe.foli21 = d21paghe.foli21   and " & _
                                               "     e21paghe.tmov21      = d21paghe.tmov21   and " & _
                                               "     e21paghe.fech21      = d21paghe.fech21   AND " & _
                                               "     TRIM(e21paghe.esta21) <> 'S'             AND " & _
                                               "     cgas21                <>''               and " & _
                                               "     cgas21='"& pCtaGas &"' "   & _
                                               "  group by cgas21                        "


                                         '      "         d21paghe.foli21 =  e21paghe.foli21  AND " & _
                                         '      "         d21paghe.tmov21 =  e21paghe.tmov21  AND " & _
                                         '      "         e21paghe.esta21 <> 'S'    " & _
                                         '      "  group by d21paghe.refe21  "



               'response.Write(StrQryConcppH& chr(13) & chr(10))
               Set RsConcppHdesg = Server.CreateObject("ADODB.Recordset")
			         RsConcppHdesg.ActiveConnection = MM_EXTRANET_STRING
			         RsConcppHdesg.Source = StrQryConcppH
			         RsConcppHdesg.CursorType = 0
			         RsConcppHdesg.CursorLocation = 2
			         RsConcppHdesg.LockType = 1
			         RsConcppHdesg.Open()
               inti=0
			         if not RsConcppHdesg.eof then
                  for inti=0 to (IntLonPH-1)
                    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& RsConcppHdesg.Fields.Item("var"&inti).Value &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
                  next
               else
                  'Response.Write("No tiene Registros")
                  for inti=0 to (IntLonPH-1)
                    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& "0" &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
                  next
               end if
               RsConcppHdesg.close
               set RsConcppHdesg = Nothing
            else
               for inti=0 to (IntLonPH-1)
                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& "0" &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
               next
            end if



              'While not RsConcppHdesg.eof
              '   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& RsConcppHdesg.Fields.Item("var"&inti).Value &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
              '   inti=inti+1
              '   RsConcppHdesg.movenext
              'wend
        '  select
        '             sum(if(conc21=1,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)  ),0 )      ) as impuesto,
        '              sum(if(conc21=2,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Maniobras,
        '              sum(if(conc21=3,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Almacenajes,
        '              sum(if(conc21=5,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as FleteTerrestre,
        '              sum(if(conc21=6,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as DemorasPorContenedor,
        '              sum(if(conc21=16,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as fumigaciones,
        '              sum(if(conc21=41,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as pagoLiberacion,
        '              sum(if(conc21=70,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as LimpiezaContenedores,
        '              sum(if(conc21=86,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ManiobrasylimpiezaCont,
        '              sum(if(conc21=100,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Revalidacion,
        '              sum(if(conc21=107,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ReparacionCont,
        '              sum(if(conc21=111,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ManiobrasyAlmacenajes,
        '              sum(if(conc21=130,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ReparacionContenedor,
        '              sum(if(conc21=151,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Prevalidacion,
        '              sum(if(conc21=153,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ProcesamientoElectronicoDatos,
        '              sum(if(conc21=181,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as maniobrasyMuellajes,
        '              sum(if(conc21=183,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as almacenajemaniobrasymuellajes
        ' from  d21paghe, e21paghe
        ' where  d21paghe.refe21               =  'RKU06-03322' and
        '             YEAR(d21paghe.fech21)    =  YEAR(e21paghe.fech21) AND
        '             d21paghe.foli21                  =  e21paghe.foli21 AND
        '             d21paghe.tmov21               =  e21paghe.tmov21 and
        '             e21paghe.esta21                <>  'S'
        '      group by d21paghe.refe21


		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"repcli18")&"</font></td>" & chr(13) & chr(10) 'Contacto
		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"rfccli18")&"</font></td>" & chr(13) & chr(10) 'R.F.C. del Cliente
            strHTML = strHTML&"</tr>"& chr(13) & chr(10)
            intContFracc=intContFracc + 1
            intContReg=intContReg + 1
		        RsRep2.movenext

   		   Wend

      end if
      RsRep2.close
      Set RsRep2=Nothing
      'Regresa el HTML del reporte
      DespliegaRepDesgRef=strHTML
end function







else
	select case strCodError
    case "1"
	   strMenjError = "Campo en Blanco Especifique!.."
	case "5","6"
	   strMenjError = "Fechas Erroneas, Verifique!"
	case "7"
	   strMenjError = "Registros No Encontrados!"
	end select
	%>
	<table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
	  <tr>
		<td><%=strMenjError%></td>
	  </tr>
	</table>
<br>
<%end if%>



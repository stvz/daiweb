<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->



<% On Error Resume Next %>


<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
oficina_zego=Session("GAduana")
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut=100000

dim desconsolidacion,maniobras,montacargas,fleteterrestre, ImportefleteInt
ImportefleteInt=0
desconsolidacion=0
maniobras=0
montacargas=0
fleteterrestre=15
'Response.Write(oficina_zego)
'Response.End
if (oficina_zego = "VER")then
 fleteterrestre="15"
 desconsolidacion="85"
 maniobras="2"
 montacargas="171"
else
  if (oficina_zego = "MEX")then
'    fleteterrestre=15 'Es un servicio complementario hasta donde se,es Servicio de Operacion
    desconsolidacion=6
    maniobras=2
    montacargas=11
  else
    if (oficina_zego = "MAN")then
	   fleteterrestre=5
       desconsolidacion=18
       maniobras=181
       montacargas=195
	else
	  if (oficina_zego = "TAM")then
	    fleteterrestre=15
        desconsolidacion=85
		maniobras=2
		montacargas=224
	  else
        if (oficina_zego = "LAR")then
  		fleteterrestre=2 '15
        desconsolidacion=109 '210
		maniobras=54'81
		montacargas=135 'ok
	    end if
	  end if
	end if
  end if
end if

strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
permi2 = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")



if not permi = "" then
  permi = "  and (" & permi & ") "
end if

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
   permi = " AND cvecli01 =" & strFiltroCliente
   permi2 = " AND B.clie31 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
   permi2 = ""
end if


strDateIni=""
strDateFin=""
strTipoPedimento= ""
strCodError = "0"

strDateIni=trim(request.Form("txtDateIni"))
strDateFin=trim(request.Form("txtDateFin"))
strTipoPedimento=trim(request.Form("rbnTipoDate"))

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
strTipoFiltro=trim(request.Form("TipoFiltro"))
strDescripcion=trim(request.Form("txtDescripcion"))
strDateIni2=trim(request.Form("txtDateIni2"))
strDateFin2=trim(request.Form("txtDateFin2"))
strTipoPedimento2=trim(request.Form("rbnTipoDate2"))


tmpTipo = ""
if strTipoFiltro  = "Fechas" then
   if strTipoPedimento  = "1" then
     tmpTipo = "IMPORTACION"
     strSQL = "SELECT adusec01,ifnull(valdol01,0) as valdol01,tipopr01,ifnull(valmer01,0) as valmer01,ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) as p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, fletes01, segros01, ifnull(cvepvc01,'0') as cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01,anexol01,incble01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !=''  order by refcia01"
', ssprov22
   end if
   if strTipoPedimento  = "2" then
     tmpTipo = "EXPORTACION"
     strSQL = "SELECT adusec01,ifnull(valdol01,0) as valdol01,tipopr01,ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) AS p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, ifnull(fletes01,0) as fletes01, ifnull(segros01,0) as segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01,anexol01,incble01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
   end if
else
   if strTipoFiltro  = "Descripcion" then
    if strTipoPedimento2  = "1" then
     tmpTipo = "IMPORTACION"
     strSQL = "SELECT distinct adusec01,ifnull(valdol01,0) as valdol01,tipopr01,ifnull(valmer01,0) as valmer01,ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) as p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, fletes01, segros01, ifnull(cvepvc01,'0') as cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01,anexol01,incble01 FROM ssdagi01,ssfrac02 WHERE refcia01 = refcia02 and d_mer102 like '%" & strDescripcion & "%' and fecpag01 >='"&FormatoFechaInv(strDateIni2)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
    ' Response.Write(strSQL)
 '    Response.End
   end if
   if strTipoPedimento2  = "2" then
     tmpTipo = "EXPORTACION"
     strSQL = "SELECT distinct adusec01,ifnull(valdol01,0) as valdol01,tipopr01,ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) AS p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, ifnull(fletes01,0) as fletes01, ifnull(segros01,0) as segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01,anexol01,incble01 FROM ssdage01,ssfrac02 WHERE refcia01 = refcia02 and d_mer102 like '%" & strDescripcion & "%' and fecpag01 >='"&FormatoFechaInv(strDateIni2)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
   end if
   end if
end if




if not trim(strSQL)="" then
		Set RsRep = Server.CreateObject("ADODB.Recordset")
		RsRep.ActiveConnection = MM_EXTRANET_STRING
     if err.number <> 0 then %>
    	<p class="ResaltadoAzul"><%RESPONSE.Write("error = " & eRR.Description )%></p>
      <% Response.End
    end if
		RsRep.Source = strSQL
		RsRep.CursorType = 0
		RsRep.CursorLocation = 2
		RsRep.LockType = 1
		RsRep.Open()



    'Response.Write(strSQL)
    'Response.End



		'Se captura el Flete Internacional

	'ImportefleteInt=cdbl(RsRep.Fields.Item("fletes01").Value)


	if not RsRep.eof then

  ' Comienza el HTML, se pintan los titulos de las columnas
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE DESGLOSADO DE REFERENCIAS DE " & tmpTipo & " </p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>" & strDateIni & " al " & strDateFin & "</p></font></strong>"
     strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	   strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)

		Hd1 = "Cliente"
	     if strTipoUsuario = MM_Cod_Cliente_Division then
	       Hd1 = "Division"
		 end if
		 if strTipoUsuario = MM_Cod_Admon or strTipoUsuario = MM_Cod_Ejecutivo_Grupo then
	       Hd1 = "Cliente"
		 end if



		strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de Pedimento</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""200"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & Hd1 & "</td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & "Division Ref" & "</td>" & chr(13) & chr(10)  '******** 21/01/2009 para Invista

		strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No.Pitex</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observaciones</font></td>" & chr(13) & chr(10)
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
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Despacho</td>" & chr(13) & chr(10)
    StrHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Diferencia ETA/DSP</td>" & chr(13) & chr(10)
    StrHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia_cliente</td>" & chr(13) & chr(10)
    StrHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Semaforo</td>" & chr(13) & chr(10)
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
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa de Recargos(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Recargos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuotas Compensatorias</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Efectivo</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Otros</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Identificadores</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Permisos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pagos Hechos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Servicios Complemetarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA HONORARIOS (Hon+ServComp+AdicHon)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Anticipos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Saldo de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contacto</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C. del Cliente</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Incoterms</td>" & chr(13) & chr(10)
'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TotalPagosHechos2</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Maniobras</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Desconsolidacion</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Montacargas</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Transportista (Flete Internacional)</td>" & chr(13) & chr(10)
'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No Factura (Flete internacional)</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe (Flete internacional)</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No Factura (Flete terrestre)</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe (Flete terrestre)</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ID FISCAL</td>" & chr(13) & chr(10)
strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Incrementables</td>" & chr(13) & chr(10)


   strHTML = strHTML & "</tr>"& chr(13) & chr(10)

	While NOT RsRep.EOF
    'Se asigna el nombre de la referencia
    strAduSec =""

     strRefer = RsRep.Fields.Item("refcia01").Value
     strAduSec =RsRep.Fields.Item("adusec01").Value
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
	 IVAHON=0



        ' Aqui se buscan los datos de la cuenta de gastos
'        strSQL1="select A.cgas31,B.fech31,B.suph31, B.coad31, B.csce31, B.chon31, B.piva31, B.anti31,B.sald31, (B.piva31/100)*((B.chon31+B.caho31+B.csce31)*if(coad31<>Null and coad31 > 0 ,-1,1)) as IVAHON from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"' " & permi2
'caho31 Adicionales a  Honorarios
        strSQL1="select A.cgas31,B.fech31,B.suph31, B.coad31, B.csce31, B.chon31, B.piva31, B.anti31,B.sald31, (B.piva31/100)*(B.chon31+B.caho31+B.csce31) as IVAHON from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"' " & permi2
        Set RsRep1 = Server.CreateObject("ADODB.Recordset")
        RsRep1.ActiveConnection = MM_EXTRANET_STRING

        if err.number <> 0 then %>
    	<p class="ResaltadoAzul"><%RESPONSE.Write("error = " & eRR.Description )%></p>
      <% Response.End
         end if

        RsRep1.Source = strSQL1
        RsRep1.CursorType = 0
        RsRep1.CursorLocation = 2
        RsRep1.LockType = 1

        'Response.Write(strSQL1)
        'Response.End

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
		      IVAHON=cdbl(RsRep1.Fields.Item("IVAHON").Value)
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
        response.Write(strHTML)
        strHTML = ""
   RsRep.movenext
   Wend
   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   end if
   RsRep.close
   Set RsRep = Nothing
'Se pinta todo el HTML formado
   response.Write(strHTML)
   if strHTML = "" then
      strHTML = "NO EXISTEN REGISTROS"
      response.Write(strHTML)
   end if
 else
   strHTML = "NO EXISTEN REGISTROS"
   response.Write(strHTML)
end if

'Funcion que va elaborando el reporte desglosado de referencias y devuelve el HTML
function DespliegaRepDesgRef(pRefer, pCtaGas)
'Sus parametros son
      'pRefer    Referencia
      'pCtaGas  Cuenta de Gastos
       dim dblEfectivo
       dim dblOtros

      strSQL3="select ifnull(fpagoi36,'0') as fpagoi36 ,ifnull(import36,0) as import36 from sscont36 where refcia36='"&pRefer&"' and adusec36='" & strAduSec & "'"

			Set RsRep3 = Server.CreateObject("ADODB.Recordset")
			RsRep3.ActiveConnection = MM_EXTRANET_STRING
			RsRep3.Source = strSQL3
			RsRep3.CursorType = 0
			RsRep3.CursorLocation = 2
			RsRep3.LockType = 1

      'Response.Write(strSQL3)
      'Response.End

			RsRep3.Open()



      if not RsRep3.eof then

      ' Aqui se obtienen los campos de Suma de Efectivo y Suma de Otros
        dblEfectivo=0
        dblOtros=0
         while not RsRep3.eof
         if RsRep3.Fields.Item("fpagoi36").Value <> "" then
          If cdbl(RsRep3.Fields.Item("fpagoi36").Value)=0 then
                dblEfectivo=dblEfectivo+cdbl(RsRep3.Fields.Item("import36").Value)  'Sumamos el efectivo para una referencia
          else
                dblOtros=dblOtros+cdbl(RsRep3.Fields.Item("import36").Value)         'Sumamos los otros conceptos para una referencia
          end if
         end if
			 RsRep3.movenext
         wend
         RsRep3.close
         Set RsRep3=Nothing
      end if
      strDivi = ""
        frag = 0
        Observaciones = RsRep.Fields.Item("anexol01").Value

         if RsRep.Fields.Item("cveadu01").Value = "80" and RsRep.Fields.Item("cvecli01").Value = 40 then
            frag = clng(InStr( Observaciones ,"FRAGANCIAS"))
            if frag > 0 then
              strDivi= "FRAGANCIAS"
            else
                 strDivi= "SABORES"
            end if
         end if

      'Aqui se obtienen las fracciones por referencia
			strSQL2="select ordfra02,fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,ifnull(vmerme02,0) as vmerme02,vaduan02,ifnull(tasadv02,0) as tasadv02,ifnull(p_adv102,0) as p_adv102,ifnull(i_adv102,0) as i_adv102,ifnull(i_iva102,0) as i_iva102,ifnull(i_adv102,0) as i_adv102,ifnull(i_adv202,0) as i_adv202,i_cc0102,ifnull(i_cc0202,0) as i_cc0202,paiori02,paiscv02 from ssfrac02 where refcia02='"&pRefer&"' and adusec02='" & strAduSec &"'"
      if strTipoFiltro  = "Descripcion" then
         strSQL2="select ordfra02,fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,ifnull(vmerme02,0) as vmerme02,vaduan02,ifnull(tasadv02,0) as tasadv02,ifnull(p_adv102,0) as p_adv102,ifnull(i_adv102,0) as i_adv102,ifnull(i_iva102,0) as i_iva102,ifnull(i_adv102,0) as i_adv102,ifnull(i_adv202,0) as i_adv202,i_cc0102,ifnull(i_cc0202,0) as i_cc0202,paiori02,paiscv02 from ssfrac02 where refcia02='"&pRefer& "' and d_mer102 like '%" & strDescripcion & "%'  and adusec02='" & strAduSec &"'"

      end if
'Response.Write(strSQL2)
'Response.End
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

         if RsRep.Fields.Item("cveadu01").Value = "80" and RsRep.Fields.Item("cvecli01").Value = 40 then
            cmp1 = strDivi
         end if

          strSQL= "select REFCIA12,ordfra12,cveide12,tipoid12,numper12 from ssipar12 where ordfra12 =" &  RsRep2.Fields.Item("ordfra02").Value & " and refcia12='" &  pRefer & "' and adusec12='" & strAduSec &"'"
        Set RsIdent = Server.CreateObject("ADODB.Recordset")
			  RsIdent.ActiveConnection = MM_EXTRANET_STRING
			  RsIdent.Source = strSQL
			  RsIdent.CursorType = 0
			  RsIdent.CursorLocation = 2
			  RsIdent.LockType = 1
			  RsIdent.Open()
        strIdentificadores=""
        strNumPermisos =""
			  if not RsIdent.eof then
           While not RsIdent.eof
           strIdentificadores = strIdentificadores  & "  " & RsIdent.Fields.Item("cveide12").Value  & "  "
           if RsIdent.Fields.Item("tipoid12").Value = "1" then
              strNumPermisos = strNumPermisos & "  " & RsIdent.Fields.Item("numper12").Value  & "  "
           end if
           RsIdent.movenext
           wend
        end if
        RsIdent.close
        set RsIdent = Nothing


        '**********************************************************************************************************************************************************************************************************************************
        strSQL= "select fdsp01, rcli01, alea01, nomdiv01 from c01refer where refe01 ='" &  pRefer & "'"
        Set RsReferencia = Server.CreateObject("ADODB.Recordset")
			  RsReferencia.ActiveConnection = MM_EXTRANET_STRING
			  RsReferencia.Source = strSQL
			  RsReferencia.CursorType = 0
			  RsReferencia.CursorLocation = 2
			  RsReferencia.LockType = 1
			  RsReferencia.Open()

        strFechaDespacho =""
        strReferenciaCliente=""
        strSemaforo=""
        strDivRef=""

			  if not RsReferencia.eof then
           strFechaDespacho = RsReferencia.Fields.Item("fdsp01").Value
           strReferenciaCliente = RsReferencia.Fields.Item("rcli01").Value
           strSemaforo = RsReferencia.Fields.Item("alea01").Value
           strDivRef   = RsReferencia.Fields.Item("nomdiv01").Value
        end if
        RsReferencia.close
        set RsReferencia = Nothing
        '**********************************************************************************************************************************************************************************************************************************


				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cmp1&"</font></td>" & chr(13) & chr(10) 'Cliente o Division
        '***********************************************************************************************************************************************************************************************************************
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strDivRef&"</font></td>" & chr(13) & chr(10) 'Cliente o Division  ' ************************************************************************************************************
        '***********************************************************************************************************************************************************************************************************************
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaNoPrograma(pRefer,"PX")&"</font></td>" & chr(13) & chr(10) 'Esta en un programa Pitex o no
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pRefer&"</font></td>" & chr(13) & chr(10) 'Referencia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&ObservRect&"</font></td>" & chr(13) & chr(10) 'Observaciones, si es rectificado o rectificacion
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("patent01").Value&"</font></td>" & chr(13) & chr(10) 'Patente
			  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento

         if intContReg=1 then
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaNoCont(pRefer)&"</font></td>" & chr(13) & chr(10) 'Numero de Contenedores
         else
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Numero de Contenedores
         end if

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("totbul01").Value&"</font></td>" & chr(13) & chr(10) 'Total de Bultos
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveped01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de pedimento
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveadu01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de aduana

        ' aqui hay q
      strNumFacturas = ""
      strNumIncoterms= ""
	  rfcProveedor = " "

      strNumFacturas = RsRep.Fields.Item("desf0101").Value
'       if RsRep.Fields.Item("cveadu01").Value = "80" OR RsRep.Fields.Item("cveadu01").Value  = "24" THEN
if(oficina_zego <> "LAR")then
      strSQL4="select cvepro39,numfac39,terfac39,irspro22,nompro39 from ssfact39,ssprov22  where cvepro39 = cvepro22 and refcia39='"&pRefer&"'"
else
      strSQL4="select cvepro39,numfac39,terfac39,'S/N' as irspro22 ,nompro39  from ssfact39 where refcia39='"&pRefer&"'"
end if
			Set RsRep4 = Server.CreateObject("ADODB.Recordset")
			RsRep4.ActiveConnection = MM_EXTRANET_STRING
			RsRep4.Source = strSQL4
			RsRep4.CursorType = 0
			RsRep4.CursorLocation = 2
			RsRep4.LockType = 1

			RsRep4.Open()
rfcProveedor=RsRep4.Fields.Item("irspro22").Value

      if RsRep.Fields.Item("cveadu01").Value = "80" OR RsRep.Fields.Item("cveadu01").Value  = "24" THEN 'Cambio pbm
       strNumFacturas = ""
       strNumFacturasProv = ""
      END IF  'Cambio pbm

      if not RsRep4.eof then
         while not RsRep4.eof
            if RsRep.Fields.Item("cveadu01").Value = "80" OR RsRep.Fields.Item("cveadu01").Value  = "24" THEN 'Cambio pbm
               strNumFacturas     =  strNumFacturas & " "  & RsRep4.Fields.Item("numfac39").Value
               strNumFacturasProv = strNumFacturasProv & " "  & RsRep4.Fields.Item("nompro39").Value
            END IF  'Cambio pbm
            strNumIncoterms =strNumIncoterms & " "  & RsRep4.Fields.Item("terfac39").Value
     	      RsRep4.movenext
         wend
      end if
      RsRep4.close
      Set RsRep4=Nothing


				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strNumFacturas&"</font></td>" & chr(13) & chr(10) 'Facturas

				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nompro01").Value&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strNumFacturasProv &"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor



				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep2.Fields.Item("paiori02").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Origen/Destino
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaGuia(pRefer)&"</font></td>" & chr(13) & chr(10) 'Si tiene guia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nombar01").Value&"</font></td>" & chr(13) & chr(10) 'Nombre del barco

			if RsRep.Fields.Item("tipopr01").Value="1" then 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
					FecPed = RsRep.Fields.Item("fecent01").Value
			else
			    FecPed = RsRep.Fields.Item("fecpre01").Value
			end if


				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FecPed&"</font></td>" & chr(13) & chr(10) 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha de pago






        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFechaDespacho&"</font></td>" & chr(13) & chr(10) 'Diferencia
        if strFechaDespacho = "" then
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Diferencia
        else
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DateDiff("d", FecPed, strFechaDespacho)&"</font></td>" & chr(13) & chr(10) 'Diferencia
        end if
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strReferenciaCliente &"</font></td>" & chr(13) & chr(10) 'Referencia del cliente
        if strSemaforo = "1" then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">ROJO</font></td>" & chr(13) & chr(10) 'Semaforo
        else
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">VERDE</font></td>" & chr(13) & chr(10) 'Semaforo
        end if

				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tipcam01").Value&"</font></td>" & chr(13) & chr(10) 'Tipo de Cambio

				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep2.Fields.Item("paiscv02").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Vendedor/Comprador

        if intContReg=1 then
         if not RsRep.Fields.Item("cveadu01").Value = "24" or RsRep.Fields.Item("cveadu01").Value = "80" then
            if RsRep.Fields.Item("tipopr01").Value="1" then 'valor Factura en Dolares si es Impo
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
            end if
            if RsRep.Fields.Item("tipopr01").Value="2" then 'valor Factura en Dolares si es Expo
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valfac01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valfac01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
            end if
         else
           if RsRep.Fields.Item("tipopr01").Value="1" then 'valor Factura en Dolares si es Impo
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("valdol01").Value &"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valdol01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
           end if
           if RsRep.Fields.Item("tipopr01").Value="2" then 'valor Factura en Dolares si es Expo
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("valdol01").Value&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valdol01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
           end if
         end if

		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoSumaValoresFracc2("vaduan02", pRefer,strAduSec)&"</font></td>" & chr(13) & chr(10) 'Valor Aduana
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fletes01").Value&"</font></td>" & chr(13) & chr(10) 'fletes
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("segros01").Value&"</font></td>" & chr(13) & chr(10) 'Seguros
        else
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val. Dol. Fact.
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val.Mcia. M.N.
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Aduana
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'fletes
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Seguros
        end if

		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("fraarn02").Value&"</font></td>" & chr(13) & chr(10) 'Fraccion Arancearia
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("d_mer102").Value&"</font></td>" & chr(13) & chr(10) 'Descripcion de la Mercancia

       if intContCtas=1 then
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cancom02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
	   else
	   	 if pCtaGas="" and intContReg=1 then
		 	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cancom02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
		 end if
		 	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) ''Cant. Factura
	   end if

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("u_medc02").Value&"</font></td>" & chr(13) & chr(10) 'Cve. Unidad Fact.

        if intContCtas=1 then
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		else
		 if pCtaGas="" and intContReg=1 then
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
         end if
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		end if

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("u_medt02").Value&"</font></td>" & chr(13) & chr(10) 'Unidad Tarifa

       if intContCtas=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
       else
         if pCtaGas="" and intContReg=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
         else
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
         end if
		   end if

			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("p_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago DTA(1)

	 	if intContReg=1 then
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Importe DTA(1)
		else
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Importe DTA(1)
    end if

     strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("tasadv02").Value&"</font></td>" & chr(13) & chr(10) 'Tasa Advalorem IGI
		 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago Advalorem IGI(1)

        if intContCtas=1 then
        	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
		else
			if pCtaGas="" and intContReg=1 then
        		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		    	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
			else
          		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		  		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA(1)
        	end if
		end if

		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("t_reca01").Value&"</font></td>" & chr(13) & chr(10) 'Tasa de Recargos

      if intContCtas=1 then
         	 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&((cdbl(RsRep2.Fields.Item("i_adv102").Value) + cdbl(RsRep2.Fields.Item("i_adv202").Value))*cdbl(RsRep.Fields.Item("t_reca01").Value))/100&"</font></td>" & chr(13) & chr(10) 'Recargos
			 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("i_cc0102").Value) + cdbl(RsRep2.Fields.Item("i_cc0202").Value)&"</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
      else
	  		if pCtaGas="" and intContReg=1 then
         	 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&((cdbl(RsRep2.Fields.Item("i_adv102").Value) + cdbl(RsRep2.Fields.Item("i_adv202").Value))*cdbl(RsRep.Fields.Item("t_reca01").Value))/100&"</font></td>" & chr(13) & chr(10) 'Recargos
			 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("i_cc0102").Value) + cdbl(RsRep2.Fields.Item("i_cc0202").Value)&"</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
			else
			 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Recargos
			 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
      		end if
	  end if

        if intContReg=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblEfectivo&"</font></td>" & chr(13) & chr(10) 'Efectivo
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblOtros&"</font></td>" & chr(13) & chr(10) 'Otros
         else
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Efectivo
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Otros
        end if
       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strIdentificadores&"</font></td>" & chr(13) & chr(10) 'Identificadores
       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strNumPermisos&"</font></td>" & chr(13) & chr(10) 'Permisos

      if not pCtaGas ="" then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pCtaGas&"</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strfech31&"</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.

          if intContCG=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsuph31 + dblcoad31&"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblcsce31&"</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblchon31&"</font></td>" & chr(13) & chr(10) 'Honorarios
 	      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&IVAHON&"</font></td>" & chr(13) & chr(10) 'IVA HONORARIOS
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&(dblcsce31 + dblchon31)*(dblpiva31/100)&"</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblanti31&"</font></td>" & chr(13) & chr(10) 'Anticipos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsald31&"</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
          else
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
  	  	  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA HONORARIOS
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
          end if
          intContCG=intContCG + 1
      else
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">/ /</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
 	  	  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA HONORARIOS
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
      end if
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"repcli18")&"</font></td>" & chr(13) & chr(10) 'Contacto
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"rfccli18")&"</font></td>" & chr(13) & chr(10) 'R.F.C. del Cliente

	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strNumIncoterms&"</font></td>" & chr(13) & chr(10) 'Incoterms

Factor=(( 100*cdbl(RsRep2.Fields.Item("vmerme02").Value))/cdbl(retornaValorTotalMcia(pRefer,strAduSec)))
 if intContCtas=1 then ' si tiene mas de una cuenta de gastos solo imprimo uno pbm
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Factor*.01* cdbl(retornaTotalPagosHechos(pRefer,maniobras,"ok"))&"</font></td>" & chr(13) & chr(10) 'Maniobras
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Factor*.01* cdbl(retornaTotalPagosHechos(pRefer,desconsolidacion,"ok"))&"</font></td>" & chr(13) & chr(10) 'Desconsolidacion
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Factor*.01* cdbl(retornaTotalPagosHechos(pRefer,montacargas,"ok"))&"</font></td>" & chr(13) & chr(10) 'Montacargas
else
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Maniobras
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Desconsolidacion
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Montacargas
end if

   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&retornaCampoNavie(pRefer,oficina_zego)&"</font></td>" & chr(13) & chr(10) 'Transportista (Flete Internacional)
 if intContCtas=1 then ' si tiene mas de una cuenta de gastos solo imprimo uno pbm
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Factor*.01*cdbl(retornaCampoDagi(pRefer,"fletes01",tmpTipo))&"</font></td>" & chr(13) & chr(10) 'Importe de Factura Transportista (Flete Internacional)
 else
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Importe de Factura Transportista (Flete Internacional)
 end if


  if(oficina_zego <> "MEX")then
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&retornaFacturasPagosFletes(pRefer,fleteterrestre,oficina_zego)&"</font></td>" & chr(13) & chr(10) 'Numeros de Factura Transportista (Flete Local)
     if intContCtas=1 then ' si tiene mas de una cuenta de gastos solo imprimo uno pbm
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Factor*.01*cdbl(retornaTotalPagosHechos(pRefer,fleteterrestre,"ok"))&"</font></td>" & chr(13) & chr(10) 'Importe de Factura Transportista (Flete Local)
	 else
       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Importe de Factura Transportista (Flete Local)
     end if
   else
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&"S/N"&"</font></td>" & chr(13) & chr(10) 'No de Factura Transportista (Flete Local)
      if intContCtas=1 then ' si tiene mas de una cuenta de gastos solo imprimo uno pbm
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Factor*.01*cdbl(retornaImporteFleteLocalMexico(pRefer))&"</font></td>" & chr(13) & chr(10) 'Importe de Factura Transportista (Flete Local)
	  else
	    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Importe de Factura Transportista (Flete Local)
   	  end if
   end if


	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&rfcProveedor&"</font></td>" & chr(13) & chr(10) 'ID FISCAL
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("incble01").Value&"</font></td>" & chr(13) & chr(10) 'Incrementables


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








<%Function CampoSumaValoresFracc2(pCampo,pRefcia,padusec)
'Regresa el numero de contenedores para esa referencia
dim tmpSumaValor2
 tmpSumaValor2=0
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING
		strSQL = "SELECT sum(ifnull("&pCampo&",0)) as SumaFracc FROM ssfrac02 WHERE refcia02='" & pRefcia & "' and adusec02='" & padusec & "' group by refcia02"
    strSQL = "SELECT sum("&pCampo&") as SumaFracc FROM ssfrac02 WHERE refcia02='" & pRefcia & "' and adusec02='" & padusec & "' group by refcia02"
   ' Response.Write( MM_EXTRANET_STRING)
   ' Response.End
		RsValor1.Source = strSQL
       RsValor1.CursorType = 0
       RsValor1.CursorLocation = 2
       RsValor1.LockType = 1
       RsValor1.Open()
	if not RsValor1.EOF then
	    tmpSumaValor2 =  RsValor1.Fields.Item("SumaFracc").Value
   end if
	 RsValor1.close
	 set  RsValor1 = nothing
    CampoSumaValoresFracc2 = tmpSumaValor2
End Function

Function retornaTotalPagosHechos(pRefcia,concepto,padusec)
'Regresa el numero de contenedores para esa referencia
dim tmpSumaValor2
 tmpSumaValor2=0
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING

'    strSQL = "select if(sum(d.mont21) is not null,sum(d.mont21),0) as monto " &_
strSQL = " SELECT ifnull(sum(if(e.DEHA21='A', d.MONT21,(d.MONT21*-1))),0) as monto "&_
	" from d21paghe as d, e21paghe as e "&_
	" where e.foli21 = d.foli21 and "&_
	" e.fech21 = d.fech21 and "&_
	" (e.esta21 = 'A'  or e.esta21='E' ) and  "&_
	" e.tmov21= 'P'  and  "&_
	" d.refe21 = '"&pRefcia&"' and "&_
	" e.conc21 = "&concepto

	'RKU07-04924
		RsValor1.Source = strSQL
       RsValor1.CursorType = 0
       RsValor1.CursorLocation = 2
       RsValor1.LockType = 1
       RsValor1.Open()
	if not RsValor1.EOF  then
	    tmpSumaValor2 =  RsValor1.Fields.Item("monto").Value
   end if
	 RsValor1.close
	 set  RsValor1 = nothing
    retornaTotalPagosHechos = tmpSumaValor2
End Function


Function retornaCampoDagi(pRefcia,campo,tmpTipo)
'Regresa algun campo de ssdagi
dim tmpSumaValor2
 tmpSumaValor2=0
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING

if(tmpTipo="IMPORTACION")then
    strSQL = " SELECT  ifnull("&campo&",0) as "&campo&" FROM ssdagi01 where refcia01 = '"&pRefcia&"' "
else
   ' if(tmpTipo="EXPORTACION")then
      strSQL = " SELECT  ifnull("&campo&",0) as "&campo&" FROM ssdage01 where refcia01 = '"&pRefcia&"' "
'	else
'	   strSQL = ""
'	end if
end if
'    strSQL = " SELECT  ifnull(fletes01,0) as fletes01 FROM ssdagi01 where refcia01 = '"&pRefcia&"' "

	'RKU07-04924
		RsValor1.Source = strSQL
       RsValor1.CursorType = 0
       RsValor1.CursorLocation = 2
       RsValor1.LockType = 1

       RsValor1.Open()
	if not RsValor1.EOF  then
	    tmpSumaValor2 =  RsValor1.Fields.Item(campo).Value
	end if
	 RsValor1.close
	 set  RsValor1 = nothing
    retornaCampoDagi = tmpSumaValor2

End Function
Function retornaCampoNavie(pRefcia,oficina)
'Regresa el nombre del trasportista internacional
dim tmpSumaValor2
 tmpSumaValor2=""
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING

if(oficina = "MEX")then
   strSQL = " Select l.desc01 as nombre from c01refer as r,c01airln as l where r.refe01 = '"&pRefcia&"' and  r.cvela01 = l.cvela01 and r.cvela01 <> ''"
else
 if(oficina = "LAR")then
    strSQL = " Select 'S/N' as nombre"
  else
    strSQL = " Select n.nom01 as nombre from c01refer as r,c01navie as n where r.refe01 = '"&pRefcia&"' and  r.naim01 = n.cve01 and r.naim01 <> ''"
  end if
end if

'RKU07-04924
    RsValor1.Source = strSQL
    RsValor1.CursorType = 0
    RsValor1.CursorLocation = 2
    RsValor1.LockType = 1
    RsValor1.Open()
	if not RsValor1.EOF  then
     tmpSumaValor2 =  RsValor1.Fields.Item("nombre").Value
	end if
	 RsValor1.close
	set  RsValor1 = nothing
     retornaCampoNavie = tmpSumaValor2

End Function

Function retornaImporteFleteLocalMexico(pRefcia)
'Regresa el nombre del trasportista internacional
dim tmpSumaValor2
 tmpSumaValor2=0
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING


   strSQL = " select Mont32 as  flete "&_
	" from d32rserv "&_
	" where Ttar32 = '00033' and refe32 = '"&pRefcia&"'"

		RsValor1.Source = strSQL
       RsValor1.CursorType = 0
       RsValor1.CursorLocation = 2
       RsValor1.LockType = 1

       RsValor1.Open()
	if not RsValor1.EOF  then
	    tmpSumaValor2 =  RsValor1.Fields.Item("flete").Value
	end if
	 RsValor1.close
	 set  RsValor1 = nothing
    retornaImporteFleteLocalMexico = tmpSumaValor2
End Function

Function retornaFacturasPagosFletes(pRefcia,concepto,padusec)
'Regresa el numero de contenedores para esa referencia
dim tmpSumaValor2
 tmpSumaValor2=""
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING

    strSQL = "select d.facpro21" &_
	" from d21paghe as d, e21paghe as e "&_
	" where e.foli21 = d.foli21 and "&_
	" e.fech21 = d.fech21 and "&_
	" (e.esta21 = 'A'  or e.esta21='E' ) and  "&_
	" e.tmov21= 'P'  and  "&_
	" d.refe21 = '"&pRefcia&"' and "&_
	" e.conc21 = "&concepto

	'RKU07-04924
		RsValor1.Source = strSQL
       RsValor1.CursorType = 0
       RsValor1.CursorLocation = 2
       RsValor1.LockType = 1
       RsValor1.Open()
	if not RsValor1.EOF  then

	while not RsValor1.eof
	    tmpSumaValor2 = RsValor1.Fields.Item("facpro21").Value&" "&tmpSumaValor2
		RsValor1.movenext
	wend
   end if
	 RsValor1.close
	 set  RsValor1 = nothing
    retornaFacturasPagosFletes = tmpSumaValor2
End Function

Function retornaValorTotalMcia(pRefcia,strAduSec)
'Regresa el numero de contenedores para esa referencia
dim tmpSumaValor2
 tmpSumaValor2=0
   Set RsValor1 = Server.CreateObject("ADODB.Recordset")
	    RsValor1.ActiveConnection = MM_EXTRANET_STRING

    strSQL = " select sum(ifnull(vmerme02,0)) as vmerme02 "&_
	         " from ssfrac02 "&_
			 " where refcia02='"&pRefcia&"' and adusec02='" & strAduSec &"'"

	'RKU07-04924
		RsValor1.Source = strSQL
       RsValor1.CursorType = 0
       RsValor1.CursorLocation = 2
       RsValor1.LockType = 1
       RsValor1.Open()
	if not RsValor1.EOF  then

	while not RsValor1.eof
	    tmpSumaValor2 = RsValor1.Fields.Item("vmerme02").Value
		RsValor1.movenext
	wend
   end if
	 RsValor1.close
	 set  RsValor1 = nothing
    retornaValorTotalMcia = tmpSumaValor2
End Function

%>

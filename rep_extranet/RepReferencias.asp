<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%' On Error Resume Next %>
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment; filename=cedulapedimento.xls"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut=100000

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

tmpTipo = ""
if strTipoPedimento  = "1" then
   tmpTipo = "IMPORTACION"
   strSQL = "SELECT tipopr01,ifnull(valmer01,0) as valmer01,ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) as p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, fletes01, segros01, ifnull(cvepvc01,'0') as cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01,tsadta01,incble01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 <>'' order by refcia01"
end if
if strTipoPedimento  = "2" then
   tmpTipo = "EXPORTACION"
   strSQL = "SELECT tipopr01, ifnull(factmo01,0) as factmo01,  ifnull(p_dta101,0) as p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, fletes01, segros01,ifnull(cvepvc01,'0') as cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01,tsadta01,incble01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 <>'' order by refcia01"

end if

if not trim(strSQL)="" then
		Set RsRep = Server.CreateObject("ADODB.Recordset")
		RsRep.ActiveConnection = MM_EXTRANET_STRING
    ' if err.number <> 0 then
	%>
    	<p class="ResaltadoAzul"><%'RESPONSE.Write("error = " & eRR.Description )
		%></p>
      <% 'Response.End
    'end if
		RsRep.Source = strSQL
		RsRep.CursorType = 0
		RsRep.CursorLocation = 2
		RsRep.LockType = 1
		RsRep.Open()

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

		'strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de Pedimento</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & Hd1 & "</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No.Pitex</font></td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</font></td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observaciones</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">AA</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No. de Contenedores</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Bultos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave de Documento</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Factura</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Icoterm</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Domicilio</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Vinculacion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Origen/Destino</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Guia/B.L.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Buque</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Entrada/Presentacion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Pago</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Diferencia</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de cambio</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factor Moneda</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Vendedor/Comprador</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val. Dol. Fact.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val.Mcia. M.N.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fletes</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Seguros</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Otros Inc.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Factura</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cve. Unidad Fact.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Tarifa</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Unidad Tarifa</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Dol.(Fraccion)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Mcia.(Fraccion)</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana(Fraccion)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago DTA(1)</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa DTA(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe DTA(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa Advalorem IGI</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago Advalorem IGI(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Advalorem IGI(1)</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FormaPagoIVA</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA(1)</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Prevalidacion</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa de Recargos(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Recargos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuotas Compensatorias</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Efectivo</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Otros</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de la C.G.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pagos Hechos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Servicios Complemetarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Anticipos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Saldo de la C.G.</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contacto</td>" & chr(13) & chr(10)
		'strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C. del Cliente</td>" & chr(13) & chr(10)
    strHTML = strHTML & "</tr>"& chr(13) & chr(10)

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
        strSQL1="select A.cgas31,B.fech31,ifnull(B.suph31,0) as suph31, ifnull(B.coad31,0) as coad31, ifnull(B.csce31,0) as csce31, ifnull(B.chon31,0) as chon31, ifnull(B.piva31,0) as piva31, ifnull(B.anti31,0) as anti31,ifnull(B.sald31,0) as sald31 from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"' " & permi2

        Set RsRep1 = Server.CreateObject("ADODB.Recordset")
        RsRep1.ActiveConnection = MM_EXTRANET_STRING

   '     if err.number <> 0 then 
   %>
    	<p class="ResaltadoAzul"><%'RESPONSE.Write("error = " & eRR.Description )
		%></p>
      <%' Response.End
      '   end if

        RsRep1.Source = strSQL1
        RsRep1.CursorType = 0
        RsRep1.CursorLocation = 2
        RsRep1.LockType = 1
'        Response.Write(strSQL1)
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
%>
<%
'Funcion que va elaborando el reporte desglosado de referencias y devuelve el HTML
function DespliegaRepDesgRef(pRefer, pCtaGas)
'Sus parametros son
      'pRefer    Referencia
      'pCtaGas  Cuenta de Gastos
       dim dblEfectivo
       dim dblOtros

      strSQL3="select cveimp36,ifnull(fpagoi36,'0') as fpagoi36 ,ifnull(import36,0) as import36 from sscont36 where refcia36='"&pRefer&"'"

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
        dblPrevalida= 0
         while not RsRep3.eof
         if RsRep3.Fields.Item("fpagoi36").Value <> "" then
          If cdbl(RsRep3.Fields.Item("fpagoi36").Value)=0 then
                dblEfectivo=dblEfectivo+cdbl(RsRep3.Fields.Item("import36").Value)  'Sumamos el efectivo para una referencia
          else
                dblOtros=dblOtros+cdbl(RsRep3.Fields.Item("import36").Value)         'Sumamos los otros conceptos para una referencia
          end if
          if RsRep3.Fields.Item("cveimp36").Value = "15" then
              dblPrevalida= dblPrevalida + cdbl(RsRep3.Fields.Item("import36").Value)
          end if
         end if

			 RsRep3.movenext
         wend
         RsRep3.close
         Set RsRep3=Nothing
      end if
	  
      'Aqui se obtienen las fracciones por referencia
			strSQL2="select fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,ifnull(vmerme02,0) as vmerme02,vaduan02,ifnull(tasadv02,0) as tasadv02,ifnull(p_adv102,0) as p_adv102,ifnull(i_adv102,0) as i_adv102,p_iva102,ifnull(i_iva102,0) as i_iva102,ifnull(i_adv102,0) as i_adv102,ifnull(i_adv202,0) as i_adv202,i_cc0102,ifnull(i_cc0202,0) as i_cc0202,paiscv02 from ssfrac02 where refcia02='"&pRefer&"'"

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
        'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&TipoOper(RsRep.Fields.Item("tipopr01").Value)&"&nbsp;</font></td>" & chr(13) & chr(10) 'Tipo de Pedimento

				cmp1 = ""
				if strTipoUsuario = MM_Cod_Cliente_Division then 'Para clientes con division
				   cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"division18")
				else
				   cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"nomcli18")
				end if

				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cmp1&"</font></td>" & chr(13) & chr(10) 'Cliente o Division
				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaNoPrograma(pRefer,"PX")&"</font></td>" & chr(13) & chr(10) 'Esta en un programa Pitex o no
				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pRefer&"</font></td>" & chr(13) & chr(10) 'Referencia
				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&ObservRect&"</font></td>" & chr(13) & chr(10) 'Observaciones, si es rectificado o rectificacion
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("patent01").Value&"</font></td>" & chr(13) & chr(10) 'Patente
        strAA =""
        if RsRep.Fields.Item("patent01").Value = "3210" then
           strAA ="LIC.ROLANDO REYES KURI"
        end if
        if RsRep.Fields.Item("patent01").Value = "3921" then
           strAA ="LUIS DE LA CRUZ REYES"
        end if
        if RsRep.Fields.Item("patent01").Value = "3857" then
           strAA ="RAFAEL MENDOZA DIAZ BARRIGA"
        end if
        if RsRep.Fields.Item("patent01").Value = "3407" then
           strAA ="YOLANDA LEYVA SALAZAR"
        end if
        if RsRep.Fields.Item("patent01").Value = "3883" then
           strAA ="MA SUSANA DE GPE FRICKE URQUIOLA"
        end if
        if RsRep.Fields.Item("patent01").Value = "3044" then
           strAA ="CARLOS HUMBERTO ZESATI ANDRADE"
        end if
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strAA&"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
			  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento

         'if intContReg=1 then
         '   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaNoCont(pRefer)&"</font></td>" & chr(13) & chr(10) 'Numero de Contenedores
         'else
         '   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Numero de Contenedores
         'end if

        'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("totbul01").Value&"</font></td>" & chr(13) & chr(10) 'Total de Bultos
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveped01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de pedimento
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveadu01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de aduana


		
        Set RsProv = Server.CreateObject("ADODB.Recordset")
			  RsProv.ActiveConnection = MM_EXTRANET_STRING
        'STRSQLTEMP="select numfac39,ifnull(fecfac39,'00/00/0000') as fecfac39,terfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39 from ssfact39 where refcia39 ='"&pRefer&"'  GROUP BY cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39"
			STRSQLTEMP="select numfac39,fecfac39, cast(fecfac39 as char) as fecfac399, terfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39 from ssfact39 where refcia39 ='"&pRefer&"'  GROUP BY cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39"
		
		RsProv.Source = STRSQLTEMP
			  RsProv.CursorType = 0
		  	RsProv.CursorLocation = 2
			  RsProv.LockType = 1
        strDomicilio = ""
        strVinculacion = ""
        strproveedor = ""
        strFacturasProveedor=""
        strFechaFacProveedor = ""
		
		
		
		
		
        strIcoterm=""
        intcuentafacturas = 0
			  RsProv.Open()
        if not RsProv.eof then
				
          while not RsProv.eof
		  
		  if intcuentafacturas = 0 then
            separador = ""
          else
            separador = " .- "
          end if
          strproveedor = RsProv.Fields.Item("nompro39").value
          strDomicilio = RsProv.Fields.Item("dompro39").value
          strIcoterm = RsProv.Fields.Item("terfac39").value
          strFacturasProveedor=strFacturasProveedor & separador  & RsProv.Fields.Item("numfac39").value
        		
          strFechaFacProveedor = strFechaFacProveedor & separador & RsProv.Fields.Item("fecfac399").value
					 
		
          if RsProv.Fields.Item("vincul39").value = "1" then
          strVinculacion = "SI"
          else
          strVinculacion = "NO"
          end if
          intcuentafacturas = intcuentafacturas + 1
           RsProv.movenext
          wend
        end if
        RsProv.close
        set RsProv = Nothing
		
		

        Set RsProv = Server.CreateObject("ADODB.Recordset")
			  RsProv.ActiveConnection = MM_EXTRANET_STRING
        'STRSQLTEMP="select numfac39,ifnull(fecfac39,'00/00/0000') as fecfac39,terfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39 from ssfact39 where refcia39 ='"&pRefer&"' GROUP BY numfac39,fecfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39"
		STRSQLTEMP="select numfac39,fecfac39, cast(fecfac39 as char) as fecfac399, terfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39 from ssfact39 where refcia39 ='"&pRefer&"' GROUP BY numfac39,fecfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39"	  
			  RsProv.Source = STRSQLTEMP
			  RsProv.CursorType = 0
		  	RsProv.CursorLocation = 2
			  RsProv.LockType = 1
        strFacturasProveedor=""
        strFechaFacProveedor = ""
        intcuentafacturas = 0
			  RsProv.Open()
		
		
		
        if not RsProv.eof then
          while not RsProv.eof
          if intcuentafacturas = 0 then
            separador = ""
          else
            separador = " .- "
          end if
          
		  strFacturasProveedor=strFacturasProveedor & separador  & RsProv.Fields.Item("numfac39").value
        	
          strFechaFacProveedor = strFechaFacProveedor & separador & RsProv.Fields.Item("fecfac399").value
        
          intcuentafacturas = intcuentafacturas + 1
           RsProv.movenext
          wend
		  
        end if
		
		
        RsProv.close
        set RsProv = Nothing
		
		
		
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFacturasProveedor&"</font></td>" & chr(13) & chr(10) 'Facturas
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFechaFacProveedor&"</font></td>" & chr(13) & chr(10) 'Facturas
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strIcoterm &"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strproveedor&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strDomicilio&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strVinculacion &"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor

				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep.Fields.Item("cvepod01").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Origen/Destino
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaGuia(pRefer)&"</font></td>" & chr(13) & chr(10) 'Si tiene guia
				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nombar01").Value&"</font></td>" & chr(13) & chr(10) 'Nombre del barco

			'if RsRep.Fields.Item("tipopr01").Value="1" then 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
			'		FecPed = RsRep.Fields.Item("fecent01").Value
			'else
			'    FecPed = RsRep.Fields.Item("fecpre01").Value
			'end if


				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FecPed&"</font></td>" & chr(13) & chr(10) 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha de pago
				'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DateDiff("d", FecPed, RsRep.Fields.Item("fecpag01").Value)&"</font></td>" & chr(13) & chr(10) 'Diferencia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tipcam01").Value&"</font></td>" & chr(13) & chr(10) 'Tipo de Cambio
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Factor Moneda
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep2.Fields.Item("paiscv02").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Vendedor/Comprador

        if intContReg=1 then
          if RsRep.Fields.Item("tipopr01").Value="1" then 'valor Factura en Dolares si es Impo
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
          end if
           if RsRep.Fields.Item("tipopr01").Value="2" then 'valor Factura en Dolares si es Expo
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valfac01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
              strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valfac01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
          end if

		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoSumaValoresFracc("vaduan02", pRefer)&"</font></td>" & chr(13) & chr(10) 'Valor Aduana
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fletes01").Value&"</font></td>" & chr(13) & chr(10) 'fletes
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("segros01").Value&"</font></td>" & chr(13) & chr(10) 'Seguros
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("incble01").Value&"</font></td>" & chr(13) & chr(10) 'Seguros

        else
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val. Dol. Fact.
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val.Mcia. M.N.
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Valor Aduana
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'fletes
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Seguros
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Otros
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

        'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("u_medc02").Value&"</font></td>" & chr(13) & chr(10) 'Cve. Unidad Fact.

        'if intContCtas=1 then
			'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		'else
		 'if pCtaGas="" and intContReg=1 then
		    'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
         'end if
		    'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		'end if

        'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("u_medt02").Value&"</font></td>" & chr(13) & chr(10) 'Unidad Tarifa

       'if intContCtas=1 then
       '   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
       '   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
       '   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
       'else
         'if pCtaGas="" and intContReg=1 then
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
         'else
         ' strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
         ' strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
         ' strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
         'end if
		   'end if

			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("p_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago DTA(1)
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tsadta01").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago DTA(1)

	 	if intContReg=1 then
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Importe DTA(1)
		else
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Importe DTA(1)
    end if

     strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("tasadv02").Value&"</font></td>" & chr(13) & chr(10) 'Tasa Advalorem IGI
		 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago Advalorem IGI(1)


    if intContCtas=1 then
        	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_iva102").Value&"</font></td>" & chr(13) & chr(10) 'Forma Pago IVA
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)

		else
			if pCtaGas="" and intContReg=1 then
      		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_iva102").Value&"</font></td>" & chr(13) & chr(10) 'F Pago
		    	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
			else
       		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Forma Pago IVA
		  		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA(1)
     	end if
		end if

if intContReg=1 then
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblPrevalida&"</font></td>" & chr(13) & chr(10) 'Prevalidacion
else
 	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Prevalidacion
end if


		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("t_reca01").Value&"</font></td>" & chr(13) & chr(10) 'Tasa de Recargos

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
		  'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblOtros&"</font></td>" & chr(13) & chr(10) 'Otros
         else
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Efectivo
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Otros
        end if

      if not pCtaGas ="" then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pCtaGas&"</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strfech31&"</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.

          if intContCG=1 then
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsuph31 + dblcoad31&"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblcsce31&"</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblchon31&"</font></td>" & chr(13) & chr(10) 'Honorarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&(dblcsce31 + dblchon31)*(dblpiva31/100)&"</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblanti31&"</font></td>" & chr(13) & chr(10) 'Anticipos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsald31&"</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
          else
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
          'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
          end if
          intContCG=intContCG + 1
      else
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">/ /</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.
		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
      end if
		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"repcli18")&"</font></td>" & chr(13) & chr(10) 'Contacto
		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"rfccli18")&"</font></td>" & chr(13) & chr(10) 'R.F.C. del Cliente
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
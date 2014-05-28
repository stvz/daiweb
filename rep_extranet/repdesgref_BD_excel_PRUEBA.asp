<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
   dim AstrCtaGas()
	 dim Astrfech31()
   dim Adblsuph31()
	 dim Adblcoad31()
	 dim Adblcsce31()
	 dim Adblchon31()
	 dim Adblpiva31()
	 dim Adblanti31()
	 dim Adblsald31()

MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut=1000

strPermisos = Request.Form("Permisos")
permi = PermisoClientesconSTR(Session("GAduana"),strPermisos,"P.cvecli01")

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
   strSQL = "SELECT P.desf0101,P.tipopr01,P.valmer01,P.factmo01,P.p_dta101,P.t_reca01,P.i_dta101,P.cvecli01,P.refcia01,P.fecpag01,P.valfac01,P.fletes01,P.segros01,P.cvepvc01,P.tipcam01,P.patent01,P.numped01,P.totbul01,P.cveped01,P.cveadu01,P.nompro01,P.cvepod01,P.nombar01,P.fecent01,F.fraarn02,F.d_mer102,F.cancom02,F.u_medc02,F.cantar02,F.u_medt02,F.vmerme02,F.vaduan02,F.tasadv02,F.p_adv102,F.i_adv102,F.i_iva102,F.i_adv202,F.i_cc0102,F.i_cc0202  FROM ssfrac02 as F LEFT JOIN ssdagi01 as P ON P.refcia01 = F.refcia02 WHERE P.fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND P.fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and P.firmae01 !='' order by P.refcia01,F.fraarn02"
'      Response.Write(strSQL)
 '  Response.End

end if
if strTipoPedimento  = "2" then
   tmpTipo = "EXPORTACION"
   strSQL = "SELECT tipopr01, factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"

end if

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
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & Hd1 & "</td>" & chr(13) & chr(10)
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
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Diferencia</td>" & chr(13) & chr(10)
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
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pagos Hechos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Servicios Complemetarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Anticipos</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Saldo de la C.G.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contacto</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C. del Cliente</td>" & chr(13) & chr(10)
    strHTML = strHTML & "</tr>"& chr(13) & chr(10)

  strFraccionAnt = ""
	While NOT RsRep.EOF
    'Se asigna el nombre de la referencia
     strRefer = RsRep.Fields.Item("refcia01").Value
     strHTML = strHTML&"<tr>"& chr(13) & chr(10)
     dim intContReg
     dim intContCtas
     dim strCtaGas
  	 dim intContCG

	 intContCtas=0
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

     if strFraccionAnt = "" then
      intContReg=1
     end if
     if strFraccionAnt = RsRep.Fields.Item("fraarn02").Value then
         intContReg=2
     end if

        ' Aqui se buscan los datos de la cuenta de gastos
        strSQL1="select A.cgas31,B.fech31,B.suph31, B.coad31, B.csce31, B.chon31, B.piva31, B.anti31,B.sald31 from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"'"

        Set RsRep1 = Server.CreateObject("ADODB.Recordset")
        RsRep1.ActiveConnection = MM_EXTRANET_STRING
        RsRep1.Source = strSQL1
        RsRep1.CursorType = 0
        RsRep1.CursorLocation = 2
        RsRep1.LockType = 1
        RsRep1.Open()
        redim AstrCtaGas(1)
	      redim Astrfech31(1)
        redim Adblsuph31(1)
	      redim Adblcoad31(1)
	      redim Adblcsce31(1)
	      redim Adblchon31(1)
	      redim Adblpiva31(1)
	      redim Adblanti31(1)
	      redim Adblsald31(1)
        while not RsRep1.eof
          intContCG=1
          redim preserve  AstrCtaGas(intContCtas+1)
          redim preserve  Astrfech31(intContCtas+1)
          redim preserve  Adblsuph31(intContCtas+1)
          redim preserve  Adblcoad31(intContCtas+1)
          redim preserve  Adblcsce31(intContCtas+1)
          redim preserve  Adblchon31(intContCtas+1)
          redim preserve  Adblpiva31(intContCtas+1)
          redim preserve  Adblanti31(intContCtas+1)
          redim preserve  Adblsald31(intContCtas+1)
          strCtaGas = RsRep1.Fields.Item("cgas31").Value
          AstrCtaGas(intContCtas)=RsRep1.Fields.Item("cgas31").Value
          Astrfech31(intContCtas)=RsRep1.Fields.Item("fech31").Value
          Adblsuph31(intContCtas)=cdbl(RsRep1.Fields.Item("suph31").Value)
          Adblcoad31(intContCtas)=cdbl(RsRep1.Fields.Item("coad31").Value)
          Adblcsce31(intContCtas)=cdbl(RsRep1.Fields.Item("csce31").Value)
          Adblchon31(intContCtas)=cdbl(RsRep1.Fields.Item("chon31").Value)
          Adblpiva31(intContCtas)=cdbl(RsRep1.Fields.Item("piva31").Value)
          Adblanti31(intContCtas)=cdbl(RsRep1.Fields.Item("anti31").Value)
          Adblsald31(intContCtas)=cdbl(RsRep1.Fields.Item("sald31").Value)
          intContCtas=intContCtas + 1
          RsRep1.movenext
        wend
		    RsRep1.close
        Set RsRep1=Nothing
        if intContCtas = 0 then
            strHTML=DespliegaRepDesgRef(strRefer, strCtaGas)
        else
        'Si tiene varias cuentas de gastos se repiten los datos de la referencia y se despliegan los distintos datos de las diferentes cuentas de gastos
         for t=1 to intContCtas
           strCtaGas=AstrCtaGas(t)
	         strfech31=Astrfech31(t)
	         dblsuph31=Adblsuph31(t)
	         dblcoad31=Adblcoad31(t)
	         dblcsce31=Adblcsce31(t)
	         dblchon31=Adblchon31(t)
	         dblpiva31=Adblpiva31(t)
	         dblanti31=Adblanti31(t)
	         dblsald31=Adblsald31(t)
           strHTML=DespliegaRepDesgRef(strRefer, strCtaGas)
         next

        end if
        strHTML = strHTML&"</tr>"& chr(13) & chr(10)
        response.Write(strHTML)
        strHTML = ""
        strFraccionAnt = RsRep.Fields.Item("fraarn02").Value

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
<%end if

%>
<%
'Funcion que va elaborando el reporte desglosado de referencias y devuelve el HTML
function DespliegaRepDesgRef(pRefer, pCtaGas)
'Sus parametros son
      'pRefer    Referencia
      'pCtaGas  Cuenta de Gastos
       dim dblEfectivo
       dim dblOtros

      strSQL3="select fpagoi36, import36 from sscont36 where refcia36='"&pRefer&"'"

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

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&TipoOper(RsRep.Fields.Item("tipopr01").Value)&"&nbsp;</font></td>" & chr(13) & chr(10) 'Tipo de Pedimento

				cmp1 = ""
				if strTipoUsuario = MM_Cod_Cliente_Division then 'Para clientes con division
				   cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"division18")
				else
				   cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"nomcli18")
				end if

				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cmp1&"</font></td>" & chr(13) & chr(10) 'Cliente o Division
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
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("desf0101").Value&"</font></td>" & chr(13) & chr(10) 'Facturas
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nompro01").Value&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep.Fields.Item("cvepod01").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Origen/Destino
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaGuia(pRefer)&"</font></td>" & chr(13) & chr(10) 'Si tiene guia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nombar01").Value&"</font></td>" & chr(13) & chr(10) 'Nombre del barco

		  	if RsRep.Fields.Item("tipopr01").Value="1" then 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
		 			FecPed = RsRep.Fields.Item("fecent01").Value
			  else
			    FecPed = RsRep.Fields.Item("fecpre01").Value
			  end if

				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FecPed&"</font></td>" & chr(13) & chr(10) 'Fecha de Entrada o Fecha de presentacion segun el tipo de referencia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha de pago
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DateDiff("d", FecPed, RsRep.Fields.Item("fecpag01").Value)&"</font></td>" & chr(13) & chr(10) 'Diferencia
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tipcam01").Value&"</font></td>" & chr(13) & chr(10) 'Tipo de Cambio
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep.Fields.Item("cvepvc01").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Vendedor/Comprador

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
        else
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val. Dol. Fact.
		      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Val.Mcia. M.N.
		      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Valor Aduana
		      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'fletes
		      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Seguros
        end if

		   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fraarn02").Value&"</font></td>" & chr(13) & chr(10) 'Fraccion Arancearia
		   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("d_mer102").Value&"</font></td>" & chr(13) & chr(10) 'Descripcion de la Mercancia

        if intContCtas=1 then
			    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cancom02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
	      else
	   	    if pCtaGas="" and intContReg=1 then
		 	       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cancom02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
		      end if
		 	    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) ''Cant. Factura
	      end if

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("u_medc02").Value&"</font></td>" & chr(13) & chr(10) 'Cve. Unidad Fact.

        if intContCtas=1 then
			    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		    else
		      if pCtaGas="" and intContReg=1 then
		        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cantar02").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
          end if
		      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Cant. Tarifa
		    end if

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("u_medt02").Value&"</font></td>" & chr(13) & chr(10) 'Unidad Tarifa

        if intContCtas=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
        else
         if pCtaGas="" and intContReg=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("vmerme02").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Valor Dol.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("vmerme02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Mcia.(Fraccion)
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("vaduan02").Value&"</font></td>" & chr(13) & chr(10) 'Valor Aduana(Fraccion)
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

       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tasadv02").Value&"</font></td>" & chr(13) & chr(10) 'Tasa Advalorem IGI
		   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("p_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago Advalorem IGI(1)

       if intContCtas=1 then
         	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
	  	    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
	   	 else
			   if pCtaGas="" and intContReg=1 then
        		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		      	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("i_iva102").Value&"</font></td>" & chr(13) & chr(10) 'IVA(1)
			   else
         		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)
		  		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA(1)
         end if
		   end if

		   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("t_reca01").Value&"</font></td>" & chr(13) & chr(10) 'Tasa de Recargos
       if intContCtas=1 then
         	 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&((cdbl(RsRep.Fields.Item("i_adv102").Value) + cdbl(RsRep.Fields.Item("i_adv202").Value))*cdbl(RsRep.Fields.Item("t_reca01").Value))/100&"</font></td>" & chr(13) & chr(10) 'Recargos
	    		 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("i_cc0102").Value) + cdbl(RsRep.Fields.Item("i_cc0202").Value)&"</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
       else
	  		   if pCtaGas="" and intContReg=1 then
         	   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&((cdbl(RsRep.Fields.Item("i_adv102").Value) + cdbl(RsRep.Fields.Item("i_adv202").Value))*cdbl(RsRep.Fields.Item("t_reca01").Value))/100&"</font></td>" & chr(13) & chr(10) 'Recargos
			       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("i_cc0102").Value) + cdbl(RsRep.Fields.Item("i_cc0202").Value)&"</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
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

      if not pCtaGas ="" then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pCtaGas&"</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
		      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strfech31&"</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.
        if intContCG=1 then
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsuph31 + dblcoad31&"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
    		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblcsce31&"</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblchon31&"</font></td>" & chr(13) & chr(10) 'Honorarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&(dblcsce31 + dblchon31)*(dblpiva31/100)&"</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblanti31&"</font></td>" & chr(13) & chr(10) 'Anticipos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsald31&"</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
        else
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Pagos Hechos
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
          strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Honorarios
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
	      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Anticipos
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
      end if
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"repcli18")&"</font></td>" & chr(13) & chr(10) 'Contacto
		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"rfccli18")&"</font></td>" & chr(13) & chr(10) 'R.F.C. del Cliente
    'Regresa el HTML del reporte
   DespliegaRepDesgRef=strHTML
end function
%>
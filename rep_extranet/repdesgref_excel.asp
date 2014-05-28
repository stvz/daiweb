
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
'permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
'permi2 = PermisoClientesTabla("B",strAduana ,strPermisos,"clie31")

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
   permi = " AND i.cvecli01 =" & strFiltroCliente
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
		case "TOL"
			strOficina="tol"
	end select



tmpTipo = ""
if strTipoPedimento  = "1" then
   tmpTipo = "IMPORTACION"
   'strSQL = "SELECT tipopr01, valmer01,factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01, nomrep01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
	strSQL="Select  i.valmer01 , i.factmo01, i.p_dta101 , i.t_reca01 ,i.i_dta101 , i.cvecli01 , i.refcia01 , i.fecpag01 , i.valfac01 ,  i.fletes01, i.segros01, i.cvepvc01, i.tipcam01,i.patent01, i.numped01, i.totbul01, i.cveped01, i.cveadu01, i.desf0101," &_
		" i.nompro01, i.cvepod01, i.nombar01, i.tipopr01, i.fecent01, i.nomrep01, " & _
		" (select group_concat(distinct s.edocum39) from "&strOficina&"_extranet.ssfact39 as s	where s.refcia39 =i.refcia01  and s.adusec39 =i.adusec01 and s.patent39 =i.patent01) as Cove, (select group_concat(distinct ip.cveide11 ,'-', ip.comide11 ) from "&strOficina&"_extranet.ssiped11  as ip where ip.refcia11=i.refcia01 and ip.cveide11='ED' and ip.patent11 =i.patent01 and ip.adusec11 =i.adusec01) as 'E-document' " & _
		" from "&strOficina&"_extranet.ssdagi01  as i where i.fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND i.fecpag01 <='"&FormatoFechaInv(strDateFin)&"' "& Permi & " and i.firmae01  !='' group by i.refcia01 order by i.refcia01"
end if
if strTipoPedimento  = "2" then
   tmpTipo = "EXPORTACION"
   'strSQL = "SELECT tipopr01, factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01, nomrep01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
	strSQL="SELECT i.tipopr01, i.factmo01, i.p_dta101, i.t_reca01, i.i_dta101, i.cvecli01, i.refcia01, i.fecpag01, i.valfac01, i.fletes01, i.segros01, i.cvepvc01, i.tipcam01, i.patent01, i.numped01, i.totbul01, i.cveped01, i.cveadu01, i.desf0101, "&_
		"i.nompro01, i.cvepod01, i.nombar01,i.tipopr01, i.fecpre01, i.nomrep01," &_
		"(select group_concat(distinct s.edocum39) from "&strOficina&"_extranet.ssfact39 as s	where s.refcia39 =i.refcia01 ) as Cove, (select group_concat(distinct ip.cveide11 ,'-', ip.comide11 ) from "&strOficina&"_extranet.ssiped11  as ip where ip.refcia11=i.refcia01 and ip.cveide11='ED') as 'E-document' " &_ 
		"FROM "&strOficina&"_extranet.ssdage01 as i" &_ 
		" WHERE i.fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND i.fecpag01 <='"&FormatoFechaInv(strDateFin)&"'  "& Permi & " and i.firmae01 !='' group by i.refcia01 order by i.refcia01"
		'response.write strSQL
		'response.end()
end if

if not trim(strSQL)="" then
		Set RsRep = Server.CreateObject("ADODB.Recordset")
		RsRep.ActiveConnection = MM_EXTRANET_STRING
		RsRep.Source = strSQL
		RsRep.CursorType = 0
		RsRep.CursorLocation = 2
		RsRep.LockType = 1
		RsRep.Open()

    'Response.Write(strSQL)
    'Response.End

	if not RsRep.eof then
  ' Comienza el HTML, se pintan los titulos de las columnas
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE DESGLOSADO DE REFERENCIAS DE " & tmpTipo & " </p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>GRUPO REYES KURI, S.C. </p></font></strong>"
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
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Identificadores</td>" & chr(13) & chr(10)
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
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">COVE</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">E-document</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Código de Producto</td>" & chr(13) & chr(10)
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
        strSQL1="select A.cgas31,B.fech31,B.suph31, B.coad31, B.csce31, B.chon31, B.piva31, B.anti31,B.sald31 from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"' " & permi2


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




      '********************************************************************************************************************************************************************************************************************************************************
         strSQLc01refer= "select nomdiv01 from c01refer where refe01 ='" &  pRefer & "'"
         Set RsReferencia = Server.CreateObject("ADODB.Recordset")
			   RsReferencia.ActiveConnection = MM_EXTRANET_STRING
			   RsReferencia.Source = strSQLc01refer
			   RsReferencia.CursorType = 0
			   RsReferencia.CursorLocation = 2
		 	   RsReferencia.LockType = 1
		 	   RsReferencia.Open()
         strDivRef=""
		  	 if not RsReferencia.eof then
            strDivRef   = RsReferencia.Fields.Item("nomdiv01").Value
         end if
      '   RsReferencia.close
      '   set RsReferencia = Nothing
      '********************************************************************************************************************************************************************************************************************************************************


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

      'Aqui se obtienen las fracciones por referencia
			strSQL2="select ordfra02,fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,vmerme02,vaduan02,tasadv02,p_adv102,i_adv102,i_iva102,i_adv102,i_adv202,i_cc0102,i_cc0202,(select concat('''',group_concat(cpro05)) from d05artic where refe05 = '"&pRefer&"' and agru05 = ordfra02 ) as cprod from ssfrac02 where refcia02='"&pRefer&"'"

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

        strSQL= "select REFCIA12,ordfra12,cveide12 from ssipar12 where ordfra12 =" &  RsRep2.Fields.Item("ordfra02").Value & " and refcia12='" &  pRefer & "'"
        Set RsIdent = Server.CreateObject("ADODB.Recordset")
			  RsIdent.ActiveConnection = MM_EXTRANET_STRING
			  RsIdent.Source = strSQL
			  RsIdent.CursorType = 0
			  RsIdent.CursorLocation = 2
			  RsIdent.LockType = 1
			  RsIdent.Open()



        strIdentificadores=""
			  if not RsIdent.eof then
           While not RsIdent.eof
           strIdentificadores = strIdentificadores  & "  " & RsIdent.Fields.Item("cveide12").Value  & "  "
           RsIdent.movenext
           wend
        end if
        RsIdent.close
        set RsIdent = Nothing

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


        'Response.Write(RsRep.Fields.Item("valmer01").Value)
        'Response.Write(RsRep.Fields.Item("factmo01").Value)
        'Response.End

        if intContReg=1 then
           if RsRep.Fields.Item("tipopr01").Value="1" then 'valor Factura en Dolares si es Impo

               if RsRep.Fields.Item("valmer01").Value <> "" and not isnull(RsRep.Fields.Item("valmer01").Value)  then
                  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
                  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("valmer01").Value) * cdbl(RsRep.Fields.Item("factmo01").Value) * cdbl(RsRep.Fields.Item("tipcam01").Value)&"</font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
               else
                  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">  </font></td>" & chr(13) & chr(10) 'Val. Dol. Fact.
                  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">  </font></td>" & chr(13) & chr(10) 'Val.Mcia. M.N.
               end if
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


      'Response.Write("RsRep2.Fields.Item(i_adv102).Value")
      'Response.Write(RsRep2.Fields.Item("i_adv102").Value)

      'Response.Write("RsRep2.Fields.Item(i_adv202).Value")
      'Response.Write(RsRep2.Fields.Item("i_adv202").Value)
      'Response.Write("RsRep.Fields.Item(t_reca01).Value")
      'Response.Write(RsRep.Fields.Item("t_reca01").Value)
      'Response.End

      if intContCtas=1 then
          if RsRep2.Fields.Item("i_adv202").Value <> "" and not isnull(RsRep2.Fields.Item("i_adv202").Value) and RsRep.Fields.Item("t_reca01").Value <> "" and not isnull(RsRep.Fields.Item("t_reca01").Value) then
         	   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&((cdbl(RsRep2.Fields.Item("i_adv102").Value) + cdbl(RsRep2.Fields.Item("i_adv202").Value))*cdbl(RsRep.Fields.Item("t_reca01").Value))/100&"</font></td>" & chr(13) & chr(10) 'Recargos
			       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep2.Fields.Item("i_cc0102").Value) + cdbl(RsRep2.Fields.Item("i_cc0202").Value)&"</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
          else
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> </font></td>" & chr(13) & chr(10) 'Recargos
			       strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> </font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
          end if
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
		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"repcli18")&"</font></td>" & chr(13) & chr(10) 'Contacto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("nomrep01").Value&"</font></td>" & chr(13) & chr(10) 'Contacto
		'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoSSdag("RKU","ssdagi01",RsRep.Fields.Item("cvecli01").Value,"nomrep01",pRefer)&"</font></td>" & chr(13) & chr(10) 'Contacto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"rfccli18")&"</font></td>" & chr(13) & chr(10) 'R.F.C. del Cliente
		strHTML = strHTML&"<td width=""160"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("Cove").Value&"</font></td>" & chr(13) & chr(10) 'Cove
		strHTML = strHTML&"<td width=""380"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("E-document").Value&"</font></td>" & chr(13) & chr(10) 'E-document
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cprod").Value&"</font></td>" & chr(13) & chr(10) 'Codigo de producto
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
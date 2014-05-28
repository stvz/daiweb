<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% On Error Resume Next %>
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE

Response.Addheader "Content-Disposition", "attachment; filename=detalladoxmercancia.xls"
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




'Response.Write("permi=")
'Response.Write(permi)
'Response.Write("<BR>permi2=")
'Response.Write(permi2)
'Response.End



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


'Response.Write(strSQL)
'Response.End


if not trim(strSQL)="" then
		Set RsRep = Server.CreateObject("ADODB.Recordset")
		RsRep.ActiveConnection = MM_EXTRANET_STRING
     if err.number <> 0 then %>
    	<p class="ResaltadoAzul"><%RESPONSE.Write("error1 = " & eRR.Description )%></p>
      <% Response.End
    end if
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



    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TipoMercancia</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Domicilio</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Vinculacion</td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Factura</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val. Dol. Fact.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Val.Mcia. M.N.</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fletes</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Seguros</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Otros Inc.</td>" & chr(13) & chr(10)





    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">AA</font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave de Documento</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</font></td>" & chr(13) & chr(10)
    		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Pago</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Guia/B.L.</td>" & chr(13) & chr(10)


strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Vendedor/Comprador</td>" & chr(13) & chr(10)

		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Origen/Destino</td>" & chr(13) & chr(10)



		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de cambio</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factor Moneda</td>" & chr(13) & chr(10)



    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cod.Mercancia</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Desc.Mercancia</td>" & chr(13) & chr(10)



		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Orden</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Lote</td>" & chr(13) & chr(10)

		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cant. Factura</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Factura</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa Advalorem IGI</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa DTA(1)</td>" & chr(13) & chr(10)
     strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Advalorem IGI(1)</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe DTA(1)</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Prevalidacion</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuotas Compensatorias</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Recargos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago Advalorem IGI(1)</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Forma de Pago DTA(1)</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FormaPagoIVA</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA(1)</td>" & chr(13) & chr(10)



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

        if err.number <> 0 then %>
    	<p class="ResaltadoAzul"><%
RESPONSE.Write("error2 = " & eRR.Description )
response.write ("Numero de error: " & Err.number  & "<br>")
response.write ("Fuente de error:" & Err.source & "<br>")
response.write ("ref: " &  RsRep.Fields.Item("refcia01").Value  & "<br>")

		%></p>
      <%

      'Response.Write(strSQL1)
      'Response.End


	  On Error Resume Next
'	  Response.End
         end if

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



         strSQLVal="select fraarn02,ordfra02,sum(vafa05) as valorfac ,nlote05 from ssfrac02 LEFT JOIN d05artic  ON refcia02 = REFE05 AND fraarn02 = frac05 and ordfra02= agru05 where refcia02='"&pRefer&"' group by refcia02"
         sumaValorFacturaTotal=0
	   		 Set RsRepVal= Server.CreateObject("ADODB.Recordset")
			   RsRepVal.ActiveConnection = MM_EXTRANET_STRING
			   RsRepVal.Source =  strSQLVal
			   RsRepVal.CursorType = 0
			   RsRepVal.CursorLocation = 2
			   RsRepVal.LockType = 1
			   RsRepVal.Open()

			   if not RsRepVal.eof then
            sumaValorFacturaTotal=  RsRepVal.Fields.Item("valorfac").Value
         end if
         RsRepVal.close
         set RsRepVal = Nothing




      'Aqui se obtienen las fracciones por referencia
			strSQL2="select nlote05,tpmerc05,vafa05,caco05,cata05,ordfra02,cpro05,dtafpp02,desc05,fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,ifnull(vmerme02,0) as vmerme02,vaduan02,ifnull(tasadv02,0) as tasadv02,ifnull(p_adv102,0) as p_adv102,ifnull(i_adv102,0) as i_adv102,p_iva102,ifnull(i_iva102,0) as i_iva102,ifnull(i_adv102,0) as i_adv102,ifnull(i_adv202,0) as i_adv202,i_cc0102,ifnull(i_cc0202,0) as i_cc0202,paiscv02 from ssfrac02 LEFT JOIN d05artic  ON refcia02 = REFE05 AND fraarn02 = frac05 and ordfra02= agru05 where refcia02='"&pRefer&"'"
			'response.write(strSQL2)
      if Session("GAduana") = "LZR" then
      			  'strSQL2="select tpmerc05,vafa05,caco05,cata05,ordfra02,cpro05,dtafpp02,desc05,fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,ifnull(vmerme02,0) as vmerme02,vaduan02,ifnull(tasadv02,0) as tasadv02,ifnull(p_adv102,0) as p_adv102,ifnull(i_adv102,0) as i_adv102,p_iva102,ifnull(i_iva102,0) as i_iva102,ifnull(i_adv102,0) as i_adv102,ifnull(i_adv202,0) as i_adv202,i_cc0102,ifnull(i_cc0202,0) as i_cc0202,paiscv02 from ssfrac02 LEFT JOIN d05artic  ON refcia02 = REFE05 AND fraarn02 = frac05 and ordfra02= pped05 where refcia02='"&pRefer&"'"
     			  strSQL2="select nlote05,ifnull(tpmerc05,'') as tpmerc05,ifnull(vafa05,0) as vafa05,ifnull(caco05,0) as caco05,ifnull(cata05,0) as cata05,ordfra02,ifnull(cpro05,'') as cpro05,dtafpp02,ifnull(desc05,'') as desc05,fraarn02,d_mer102,cancom02,cancom02,u_medc02,cantar02,u_medt02,ifnull(vmerme02,0) as vmerme02,vaduan02,ifnull(tasadv02,0) as tasadv02,ifnull(p_adv102,0) as p_adv102,ifnull(i_adv102,0) as i_adv102,p_iva102,ifnull(i_iva102,0) as i_iva102,ifnull(i_adv102,0) as i_adv102,ifnull(i_adv202,0) as i_adv202,i_cc0102,ifnull(i_cc0202,0) as i_cc0202,paiscv02 from ssfrac02 LEFT JOIN d05artic  ON refcia02 = REFE05 AND fraarn02 = frac05 and ordfra02= pped05 where refcia02='"&pRefer&"'"

      end if

			Set RsRep2 = Server.CreateObject("ADODB.Recordset")
			RsRep2.ActiveConnection = MM_EXTRANET_STRING
			RsRep2.Source = strSQL2

			RsRep2.CursorType = 0
			RsRep2.CursorLocation = 2
			RsRep2.LockType = 1
			RsRep2.Open()

	if not RsRep2.eof then
      	While not RsRep2.eof

         strSQLVal="select nlote05, fraarn02,ordfra02,sum(vafa05) as valorfac from ssfrac02 LEFT JOIN d05artic  ON refcia02 = REFE05 AND fraarn02 = frac05 and ordfra02= agru05 where fraarn02 ='" & RsRep2.Fields.Item("fraarn02").Value &"' and ordfra02='" & RsRep2.Fields.Item("ordfra02").Value &"' and refcia02='"&pRefer&"' group by fraarn02,ordfra02"
           if Session("GAduana") = "LZR" then
      	   strSQLVal="select nlote05, fraarn02,ordfra02,sum(vafa05) as valorfac from ssfrac02 LEFT JOIN d05artic  ON refcia02 = REFE05 AND fraarn02 = frac05 and ordfra02= pped05 where fraarn02 ='" & RsRep2.Fields.Item("fraarn02").Value &"' and ordfra02='" & RsRep2.Fields.Item("ordfra02").Value &"' and refcia02='"&pRefer&"' group by fraarn02,ordfra02"
           end if


	   		 Set RsRepVal= Server.CreateObject("ADODB.Recordset")
			   RsRepVal.ActiveConnection = MM_EXTRANET_STRING
			   RsRepVal.Source =  strSQLVal
			   RsRepVal.CursorType = 0
			   RsRepVal.CursorLocation = 2
			   RsRepVal.LockType = 1
			   RsRepVal.Open()
         sumaValorFactura=0
			   if not RsRepVal.eof then
                 sumaValorFactura=  RsRepVal.Fields.Item("valorfac").Value
               end if
         RsRepVal.close
         set RsRepVal = Nothing


      'Recorremos las Fracciones
				strHTML = strHTML&"<tr>" & chr(13) & chr(10)

				cmp1 = ""
				if strTipoUsuario = MM_Cod_Cliente_Division then 'Para clientes con division
				   cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"division18")
				else
				   cmp1=CampoCliente(RsRep.Fields.Item("cvecli01").Value,"nomcli18")
				end if





        Set RsProv = Server.CreateObject("ADODB.Recordset")
			  RsProv.ActiveConnection = MM_EXTRANET_STRING
        STRSQLTEMP="select numfac39,fecfac39,terfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39 from ssfact39 where refcia39 ='"&pRefer&"' GROUP BY cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39"
			  RsProv.Source = STRSQLTEMP
			  RsProv.CursorType = 0
		  	RsProv.CursorLocation = 2
			  RsProv.LockType = 1
        strDomicilio = ""
        strVinculacion = ""
        strproveedor = ""
		Auxtpmerc = ""
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
          strFechaFacProveedor = strFechaFacProveedor & separador & RsProv.Fields.Item("fecfac39").value
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
        STRSQLTEMP="select numfac39,fecfac39,terfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39 from ssfact39 where refcia39 ='"&pRefer&"' GROUP BY numfac39,fecfac39,cvepro39,nompro39,dompro39,idfisc39,vincul39,refcia39,cvepro39,noipro39,noepro39,cp_pro39,mc_pro39,nomppr39"
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
          strFechaFacProveedor = strFechaFacProveedor & separador & RsProv.Fields.Item("fecfac39").value
          intcuentafacturas = intcuentafacturas + 1
           RsProv.movenext
          wend
        end if
        RsProv.close
        set RsProv = Nothing



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


			   if not RsRep2.eof then
				  Auxtpmerc = RsRep2.Fields.Item("tpmerc05").Value
			  else
			      c = ""
              end if


  '      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&  RsRep2.Fields.Item("tpmerc05").Value &"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Auxtpmerc&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strproveedor&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strDomicilio&"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strVinculacion &"</font></td>" & chr(13) & chr(10) 'Nombre de proveedor
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFechaFacProveedor&"</font></td>" & chr(13) & chr(10) 'Fecha Facturas
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFacturasProveedor&"</font></td>" & chr(13) & chr(10) 'Facturas
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& (RsRep2.Fields.Item("vafa05").Value * cdbl(RsRep.Fields.Item("factmo01").Value)) &"</font></td>" & chr(13) & chr(10) 'Valor Factura
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& (RsRep2.Fields.Item("vafa05").Value * cdbl(RsRep.Fields.Item("factmo01").Value)*RsRep.Fields.Item("tipcam01").Value) &"</font></td>" & chr(13) & chr(10) 'Valor Factura

        dblValAduana = 0
        dblValAduana = (RsRep2.Fields.Item("vafa05").Value * RsRep2.Fields.Item("vaduan02").Value) /  sumaValorFactura
		'response.Write(RsRep2.Fields.Item("vafa05").Value&","&RsRep.Fields.Item("fletes01").Value&","&RsRep.Fields.Item("incble01").Value&","&RsRep2.Fields.Item("i_adv102").Value&","&RsRep2.Fields.Item("i_cc0102").Value&","&RsRep.Fields.Item("t_reca01").Value)
		'response.End()


		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblValAduana&"</font></td>" & chr(13) & chr(10) 'Valor Aduana

        dblFlete=0
        dblFlete =  (RsRep2.Fields.Item("vafa05").Value * RsRep.Fields.Item("fletes01").Value)/  sumaValorFacturaTotal

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblFlete&"</font></td>" & chr(13) & chr(10) 'fletes

        dblSeguros=0
        dblSeguros =  (RsRep2.Fields.Item("vafa05").Value * RsRep.Fields.Item("segros01").Value)/  sumaValorFacturaTotal

		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblSeguros &"</font></td>" & chr(13) & chr(10) 'Seguros

        dblIncrementables=0
        dblIncrementables =  (RsRep2.Fields.Item("vafa05").Value * RsRep.Fields.Item("incble01").Value)/  sumaValorFacturaTotal

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblIncrementables&"</font></td>" & chr(13) & chr(10) 'Seguros
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strAA&"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveped01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de pedimento
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("patent01").Value&"</font></td>" & chr(13) & chr(10) 'Patente
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveadu01").Value&"</font></td>" & chr(13) & chr(10) 'Clave de aduana
			  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha de pago
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RegresaGuia(pRefer)&"</font></td>" & chr(13) & chr(10) 'Si tiene guia
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep2.Fields.Item("paiscv02").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Vendedor/Comprador
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&DescPais(RsRep.Fields.Item("cvepod01").Value)&"</font></td>" & chr(13) & chr(10) 'Pais Origen/Destino
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tipcam01").Value&"</font></td>" & chr(13) & chr(10) 'Tipo de Cambio
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&cdbl(RsRep.Fields.Item("factmo01").Value)&"</font></td>" & chr(13) & chr(10) 'Factor Moneda
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("cpro05").Value&"</font></td>" & chr(13) & chr(10) 'Fraccion Arancearia
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("desc05").Value&"</font></td>" & chr(13) & chr(10) 'Descripcion de la Mercancia
    		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("fraarn02").Value&"</font></td>" & chr(13) & chr(10) 'Fraccion Arancearia
	    	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("d_mer102").Value&"</font></td>" & chr(13) & chr(10) 'Descripcion de la Mercancia
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("ordfra02").Value&"</font></td>" & chr(13) & chr(10) 'Orden
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("nlote05").Value&"</font></td>" & chr(13) & chr(10) 'Lote 
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("caco05").Value&"</font></td>" & chr(13) & chr(10) 'Cant. Factura
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("vafa05").Value&"</font></td>" & chr(13) & chr(10) 'valor mercancia
		
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("tasadv02").Value&"</font></td>" & chr(13) & chr(10) 'Tasa Advalorem IGI
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tsadta01").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago DTA(1)

        valIGIprorrateado = 0
        valIGIprorrateado = (RsRep2.Fields.Item("vafa05").Value * RsRep2.Fields.Item("i_adv102").Value) / sumaValorFactura
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&valIGIprorrateado&"</font></td>" & chr(13) & chr(10) 'Advalorem IGI(1)

        valdtaprorrateado = 0
        valdtaprorrateado = (RsRep2.Fields.Item("vafa05").Value * RsRep2.Fields.Item("dtafpp02").Value) / sumaValorFactura
			  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& valdtaprorrateado &"</font></td>" & chr(13) & chr(10) 'Importe DTA(1)

        dblPrevalidaprorrateado=0
        dblPrevalidaprorrateado=  (RsRep2.Fields.Item("vafa05").Value * dblPrevalida) /  sumaValorFacturaTotal
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblPrevalidaprorrateado&"</font></td>" & chr(13) & chr(10) 'Prevalidacion

        dblCuotaprorrateado=0
        dblCuotaprorrateado =  (RsRep2.Fields.Item("vafa05").Value * cdbl(RsRep2.Fields.Item("i_cc0102").Value) + cdbl(RsRep2.Fields.Item("i_cc0202").Value)) /  sumaValorFacturaTotal

        dblRecargosprorrateado=0
        dblRecargosprorrateado =  (RsRep2.Fields.Item("vafa05").Value * ((cdbl(RsRep2.Fields.Item("i_adv102").Value) + cdbl(RsRep2.Fields.Item("i_adv202").Value))*cdbl(RsRep.Fields.Item("t_reca01").Value))/100) /  sumaValorFacturaTotal


        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblCuotaprorrateado  &"</font></td>" & chr(13) & chr(10) 'Cuotas Compensatorias
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblRecargosprorrateado&"</font></td>" & chr(13) & chr(10) 'Recargos
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pCtaGas&"</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_adv102").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago Advalorem IGI(1)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("p_dta101").Value&"</font></td>" & chr(13) & chr(10) 'Forma de Pago DTA(1)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep2.Fields.Item("p_iva102").Value&"</font></td>" & chr(13) & chr(10) 'F Pago

        valIVAprorrateado = 0
        valIVAprorrateado = (RsRep2.Fields.Item("vafa05").Value * RsRep2.Fields.Item("i_iva102").Value) / sumaValorFactura
		    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&valIVAprorrateado&"</font></td>" & chr(13) & chr(10) 'IVA(1)


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
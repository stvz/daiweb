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

'response.Write("Permisos="&strPermisos)

if not permi = "" then
  permi = "  and (" & permi & ") "
end if


'response.Write("Permisos="&permi)


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



if strTipoFiltro  = "Fechas" then    'Filtro por fechas

   if strTipoPedimento  = "1" then
      tmpTipo = "IMPORTACION"
      'strSQL = "SELECT tipopr01, valmer01,factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"

      strSQL = "SELECT  refcia01, " & _
               "         patent01," & _
               "         numped01," & _
               "         fecpag01," & _
               "         tipcam01," & _
               "         factmo01," & _
               "         fraarn02 as fraccion," & _
               "         ordfra02 as orden," & _
               "         vmerme02*FACTMO01 as valorComercialDolares," & _
               "         vmerme02*FACTMO01*TIPCAM01 as valorComercialMN," & _
               "         tasadv02    as tasaIGI," & _
               "         (I_adv102) as IGIMN," & _
               "         dtafpp02   as DTAMN," & _
               "         (I_iva102) AS IVAMN" & _
               " from  SSDAGI01,  SSFRAC02 " & _
               " WHERE  REFCIA02 =REFCIA01 AND " & _
               "        fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
               "        fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & _
               Permi & " and " & _
               "   firmae01 !=''  and " & _
               "LTRIM(CVEPED01) != 'R1' " & _
               "order by refcia01"

   end if
   if strTipoPedimento  = "2" then
      tmpTipo = "EXPORTACION"
      'strSQL = "SELECT tipopr01, factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
      strSQL = "SELECT  refcia01, " & _
               "         patent01," & _
               "         numped01," & _
               "         fecpag01," & _
               "         tipcam01," & _
               "         factmo01," & _
               "         fraarn02 as fraccion," & _
               "         ordfra02 as orden," & _
               "         vmerme02*FACTMO01 as valorComercialDolares," & _
               "         vmerme02*FACTMO01*TIPCAM01 as valorComercialMN," & _
               "         tasadv02    as tasaIGI," & _
               "         (I_adv102) as IGIMN," & _
               "         dtafpp02   as DTAMN," & _
               "         (I_iva102) AS IVAMN" & _
               " from  SSDAGE01,  SSFRAC02 " & _
               " WHERE  REFCIA02 =REFCIA01 AND " & _
               "        fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
               "        fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & _
               Permi & " and " & _
               "   firmae01 !=''  and " & _
               "LTRIM(CVEPED01) != 'R1' " & _
               "order by refcia01"
   end if

end if

'response.write(strSQL)
'response.end


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
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE CONTROL DE EMBARQUES DE " & tmpTipo & " </p></font></strong>"
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>GRUPO REYES KURI, S.C. </p></font></strong>"
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>" & strDateIni & " al " & strDateFin & "</p></font></strong>"
       strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	     strHTML = strHTML & "<tr bgcolor=""#009966"" align=""center"">"& chr(13) & chr(10)

		   Hd1 = "Cliente"
 	     if strTipoUsuario = MM_Cod_Cliente_Division then
	        Hd1 = "Division"
		   end if
		   if strTipoUsuario = MM_Cod_Admon or strTipoUsuario = MM_Cod_Ejecutivo_Grupo then
	        Hd1 = "Cliente"
		   end if

		   strHTML = strHTML & "<td width=""60"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente </td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""90""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Pedimento Nº  </td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""130""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Fecha Pedimento </font></td>" & chr(13) & chr(10)
		   'strHTML = strHTML & "<td width=""100""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Identificadores </font></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""75""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Permiso Nº </font></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""75""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Fracciones</font></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""60""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Partidas </font></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""140""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Valor Comercial MXP</td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""140""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Valor Comercial USD</td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""50""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> IGI % </td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""80""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> IGI MXP</td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""80""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> DTA MXP</td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""125""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Prevalidación MXP</td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""80""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> IVA MXP</td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""95""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Recargos MXP</td>" & chr(13) & chr(10)
       strHTML = strHTML & "</tr>"& chr(13) & chr(10)



	     While NOT RsRep.EOF
       'Se asigna el nombre de la referencia
           strRefer = RsRep.Fields.Item("refcia01").Value
           strOrden = RsRep.Fields.Item("orden").Value


           dblMonto=0
           strIdentificadores=" "
           strpermisos=" "
           dblRsPRV=0




             ' Calculamos el recargo en MXP
                 strSQL1= "SELECT  IF(ltrim(cveimp36)='7' OR ltrim(cveimp36)='13' ,import36,0) as monto " & _
                          " from  SSCONT36 " & _
                          " WHERE   refcia36='" &  strRefer & "' AND  " & _
                          " (ltrim(cveimp36)='7' OR ltrim(cveimp36)='13') " & _
                          " order by refcia36 "

                 'response.Write("strSQL1="&strSQL1)
                 Set RsRep1 = Server.CreateObject("ADODB.Recordset")
                 RsRep1.ActiveConnection = MM_EXTRANET_STRING
                 RsRep1.Source = strSQL1
                 RsRep1.CursorType = 0
                 RsRep1.CursorLocation = 2
                 RsRep1.LockType = 1
                 RsRep1.Open()

                 if not RsRep1.eof then
                   while not RsRep1.eof
                     dblMonto=dblMonto + RsRep1.Fields.Item("monto").Value
                     RsRep1.movenext
                   wend
                 end if
		             RsRep1.close
                 Set RsRep1=Nothing



                              ' Traemos los identificadores de cada partida
                 strSQL2= "select REFCIA12,ordfra12,cveide12,numper12 from ssipar12 where ordfra12 =" &  strOrden & " and refcia12='" &  strRefer & "'"
                 Set RsIdent = Server.CreateObject("ADODB.Recordset")
			           RsIdent.ActiveConnection = MM_EXTRANET_STRING
			           RsIdent.Source = strSQL2
			           RsIdent.CursorType = 0
			           RsIdent.CursorLocation = 2
			           RsIdent.LockType = 1
			           RsIdent.Open()

			           if not RsIdent.eof then
                    While not RsIdent.eof
                      strIdentificadores = strIdentificadores  & "  " & RsIdent.Fields.Item("cveide12").Value  & " &nbsp; "
                      strpermisos= strpermisos & "  " & RsIdent.Fields.Item("numper12").Value  & " &nbsp; "
                      RsIdent.movenext
                    wend
                 end if
                 RsIdent.close
                 set RsIdent = Nothing

                 if strpermisos="" then
                   strpermisos=" &nbsp;"
                 end if

                 if strIdentificadores="" then
                   strIdentificadores=" &nbsp;"
                 end if


            ' Calculamos el PRV
                 strSQL3= "SELECT  IF(ltrim(cveimp36)='15' ,import36,0) as PRV " & _
                          " from  SSCONT36 " & _
                          " WHERE   refcia36='" &  strRefer & "' AND  " & _
                          " ltrim(cveimp36)='15' " & _
                          " order by refcia36 "
                 Set RsPRV = Server.CreateObject("ADODB.Recordset")
			           RsPRV.ActiveConnection = MM_EXTRANET_STRING
			           RsPRV.Source = strSQL3
     	           RsPRV.CursorType = 0
			           RsPRV.CursorLocation = 2
			           RsPRV.LockType = 1
			           RsPRV.Open()

			           if not RsPRV.eof then
                   While not RsPRV.eof
                     dblRsPRV = dblRsPRV + RsPRV.Fields.Item("PRV").Value
                     RsPRV.movenext
                   wend
                 end if
                 RsPRV.close
                 set RsPRV = Nothing



           strHTML = strHTML&"<tr>" & chr(13) & chr(10)
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("patent01").Value&"&nbsp;</font></td>" & chr(13) & chr(10) 'Patente
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10) 'Pedimento Nº
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha Pedimento
				   'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strIdentificadores&"</font></td>" & chr(13) & chr(10) 'Identificadores
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strpermisos&"</font></td>" & chr(13) & chr(10) 'Permiso Nº
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fraccion").Value&"&nbsp;"&"</font></td>" & chr(13) & chr(10) 'Fracciones
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strOrden&"</font></td>" & chr(13) & chr(10) 'Partidas
			     strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("valorComercialMN").Value&"</font></td>" & chr(13) & chr(10) 'Valor Comercial MXP
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("valorComercialDolares").Value&"</font></td>" & chr(13) & chr(10) 'Valor Comercial USD
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("tasaIGI").Value&"</font></td>" & chr(13) & chr(10) 'IGI %
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("IGIMN").Value&"</font></td>" & chr(13) & chr(10) 'IGI MXP
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("DTAMN").Value&"</font></td>" & chr(13) & chr(10) 'IGI MXP
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblRsPRV&"</font></td>" & chr(13) & chr(10) 'Prevalidación MXP
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("IVAMN").Value&"</font></td>" & chr(13) & chr(10) 'IVA MXP
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblMonto&"</font></td>" & chr(13) & chr(10) 'Recargos MXP
           strHTML = strHTML&"</tr>"& chr(13) & chr(10)

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
<%end if%>
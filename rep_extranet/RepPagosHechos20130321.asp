<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment;"
 Response.ContentType = "application/vnd.ms-excel"

 Server.ScriptTimeOut=200000

 strHTML = ""

 strDate=trim(request.Form("txtDateIni"))
 strDate2 = trim(request.Form("txtDateFin"))

 if not strDate="" and not strDate2="" then


   tmpDiaFin = cstr(datepart("d",strDate))
   tmpMesFin = cstr(datepart("m",strDate))
   tmpAnioFin = cstr(datepart("yyyy",strDate))
   strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin

   tmpDiaFin2 = cstr(datepart("d",strDate2))
   tmpMesFin2 = cstr(datepart("m",strDate2))
   tmpAnioFin2 = cstr(datepart("yyyy",strDate2))
   strDateFin2 = tmpAnioFin2 & "/" &tmpMesFin2 & "/"& tmpDiaFin2


   dim con,Rsio,Rsio2


strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"clie01")

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
   permi = " AND clie01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if


   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection = MM_EXTRANET_STRING

   strSQL = "select CLIE01 as Cliente,E21PAGHE.FOLI21 as Folio,E21PAGHE.tpag21,E21PAGHE.ESTA21 as Status,E21PAGHE.TMOV21,E21PAGHE.FECH21 as FechaPago,D21PAGHE.REFE21 as Referencia, If(E21PAGHE.DEHA21='A', D21PAGHE.MONT21,(D21PAGHE.MONT21)*-1) as Importe, C21PAGHE.DESC21 as Concepto from D21PAGHE, E21PAGHE, C21PAGHE , C01REFER WHERE (C01REFER.CLIE01 = 1928) AND D21PAGHE.REFE21 = C01REFER.REFE01 AND ( E21PAGHE.FOLI21 = D21PAGHE.FOLI21 AND YEAR(E21PAGHE.FECH21) = YEAR(D21PAGHE.FECH21) AND E21PAGHE.TMOV21 = D21PAGHE.TMOV21) AND E21PAGHE.CONC21 = C21PAGHE.CLAV21 AND E21PAGHE.FECH21 >='"&strDateFin&"' and E21PAGHE.FECH21 <='"&strDateFin2&"' AND E21PAGHE.ESTA21 <> 'S' and E21PAGHE.tpag21 <> 3 AND E21PAGHE.TMOV21 = 'P'" &  permi
   'strSQL = "select clie01 as Cliente, d21paghe.cgas21 as Cuenta,e21paghe.foli21 as Folio,e21paghe.tpag21,e21paghe.esta21 as Status,e21paghe.tmov21,e21paghe.fech21 as FechaPago,d21paghe.refe21 as referencia,If(e21paghe.deha21='A', d21paghe.mont21,(d21paghe.mont21)*-1) as Importe,c21paghe.desc21 as Concepto from e21paghe,c01refer inner join d21paghe,c21paghe on (e21paghe.foli21=d21paghe.foli21 and year(e21paghe.fech21)=year(d21paghe.fech21) and e21paghe.tmov21=d21paghe.tmov21 and d21paghe.refe21 = c01refer.refe01 and e21paghe.conc21=c21paghe.clav21) and (e21paghe.fech21>='"&strDateFin&"' and e21paghe.fech21<='"&strDateFin2&"') and e21paghe.esta21 <> 'S' and e21paghe.tpag21 <> 3 and e21paghe.tmov21='P'" &  permi
   'strSQL = "select clie01 as Cliente, d21paghe.cgas21 as Cuenta,e21paghe.fech21 as FechaPago,d21paghe.refe21 as referencia,sum(If(e21paghe.deha21='A', d21paghe.mont21,(d21paghe.mont21)*-1)) as Importe,c21paghe.desc21 as Concepto from e21paghe,c01refer inner join d21paghe,c21paghe on (e21paghe.foli21=d21paghe.foli21 and year(e21paghe.fech21)=year(d21paghe.fech21) and e21paghe.tmov21=d21paghe.tmov21 and d21paghe.refe21 = c01refer.refe01 and e21paghe.conc21=c21paghe.clav21) and (e21paghe.fech21>='"&strDateFin&"' and e21paghe.fech21<='"&strDateFin2&"') and e21paghe.esta21 <> 'S' and e21paghe.tpag21 <> 3 and e21paghe.tmov21='P'" &  permi & " group by d21paghe.refe21,c21paghe.desc21  having Importe > 0 "
   'strSQL = "select CLIE01 as Cliente,D21PAGHE.* from D21PAGHE left join C01REFER on REFE21 = REFE01 WHERE (C01REFER.CLIE01 = 1928) AND (D21PAGHE.FECH21 >='"&strDateFin&"' and D21PAGHE.FECH21 <='"&strDateFin2&"') "
	
   'Response.End()
   Rsio.Source= strSQL
   Rsio.CursorType = 0
   Rsio.CursorLocation = 2
   Rsio.LockType = 1
   'Response.Write(strSQL)
   'Response.End
  Rsio.Open()

	 strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE PAGOS HECHOS</p></font></strong>"
	 strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p> Del " & strDate & " al " & strDate2 & " </p></font></strong>"
   strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	 strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de pago</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Concepto</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos" & chr(13) & chr(10)
   strHTML = strHTML & "</tr>"& chr(13) & chr(10)


   While NOT Rsio.EOF


        strHTML = strHTML&"<tr>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("FechaPago").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Referencia").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Importe").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Concepto").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Cuenta").Value&"</font></td>" & chr(13) & chr(10)
        Response.Write( strHTML )
         strHTML = ""



  Rsio.MoveNext()
  Wend

Rsio.Close()
Set Rsio = Nothing

end if
strHTML = strHTML & "</tr>"& chr(13) & chr(10)
strHTML = strHTML & "</td>"& chr(13) & chr(10)
strHTML = strHTML & "</table>"& chr(13) & chr(10)
response.Write(strHTML)
%>
